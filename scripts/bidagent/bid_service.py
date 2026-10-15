"""精简标书智能体服务层。

职责：
- 读取 / 校验企业资质结构化清单
- 文档解析（复用 services.bidding_documents.parse_document）
- 资质逐条比对（规则）与定性可行性评估
- 调用 Dify 工作流做招标解读（WF1）与初稿内容生成（WF2），未配置时回退本地启发式
- 生成 Word 初稿（复用 services.bidding_documents.build_bid_documents）

同步阻塞操作（解析、Word 生成、httpx 同步调用）由路由层通过 asyncio.to_thread
包装；本模块的 Dify 调用使用 async httpx。
"""
from __future__ import annotations

import json
import logging
import re
import sys
from pathlib import Path
from typing import Any

import httpx

from .settings import settings
from .bidding_documents import DocumentError, build_bid_documents, parse_document

logger = logging.getLogger(__name__)

STATUS_SATISFIED = "满足"
STATUS_UNVERIFIED = "待补充"
STATUS_MISSING = "不满足"
STATUS_NA = "不适用"

# 判定“资质类要求”的关键词：命中才参与资质比对；流程/承诺/报价等要求标为“不适用”，不计入门槛缺口。
_QUALIFICATION_PATTERN = re.compile(
    r"认证|许可证|营业执照|资质证书|等级证书|备案|ISO\s?\d|管理体系|"
    r"纳税信用\s?A\s?级|守合同重信用|企业信用等级|资信证明|信用中国|"
    r"一般纳税人|预包装食品|出版物经营|医疗器械|增值电信"
)

_UNVERIFIED = "待核实"
_DISCLAIMER = "本结论为基于资质的定性可行性评估，非中标概率，仅供内部参考。"

_QUALIFICATION_HINTS = ("质量管理体系", "环境管理体系", "职业健康安全", "信息安全管理", "ISO", "认证证书",
                        "许可证", "资质", "等级证书", "备案", "纳税人", "营业执照", "信用")
_REQUIREMENT_HINTS = ("须", "应当", "应", "必须", "不得", "要求", "具备", "提供", "提交", "加盖", "承诺")
_MANDATORY_HINTS = ("否决", "废标", "无效投标", "必须", "须", "应当")


class BidError(ValueError):
    """精简标书模块的可预期错误（对应 HTTP 422）。"""


# ---------- 资质清单 ----------


def resolve_manifest_path(path: str | None = None) -> Path:
    """定位资质清单：显式路径 > 配置路径 > 输入目录下的常见位置。

    后两种兜底是为 sandbox 场景准备的：`BID_MANIFEST_PATH` 可能没配，
    而用户把清单传到输入目录（如 `/workspace/input/manifest.json`）。
    """
    if path:
        return Path(path)
    configured = Path(settings.BID_MANIFEST_PATH)
    if configured.is_file():
        return configured
    input_dir = Path(settings.bid_input_dir)
    for candidate in (input_dir / "manifest.json",
                      input_dir / "qualifications" / "manifest.json"):
        if candidate.is_file():
            return candidate
    # 都没找到：把**用户该放的位置**回报出去，便于按报错直接补齐。
    # 显式配过 BID_MANIFEST_PATH 就报它（意图明确），否则报输入目录下的默认位置。
    if settings.BID_MANIFEST_PATH.strip():
        return configured
    return input_dir / "manifest.json"


def _locate_attachment(base: Path, raw: Path) -> Path | None:
    """定位资质原件/附件：先按清单里的相对路径，再兜底到输入目录下按文件名找。

    兜底同样是为 sandbox 准备的：用户把原件**平铺**上传到输入目录时，
    清单里写的「优识企业资质/xxx.jpg」这种相对路径会落空。
    """
    if raw.is_absolute():
        return raw if raw.is_file() else None
    candidate = base / raw
    if candidate.is_file():
        return candidate
    fallback = Path(settings.bid_input_dir) / raw.name
    return fallback if fallback.is_file() else None


def load_manifest(path: str | None = None) -> dict:
    """读取并最小校验资质清单。"""
    manifest_path = resolve_manifest_path(path)
    if not manifest_path.is_file():
        raise BidError(f"资质清单不存在：{manifest_path}")
    try:
        data = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise BidError(f"资质清单无法解析：{type(exc).__name__}") from exc
    items = data.get("items")
    if not isinstance(items, list):
        raise BidError("资质清单缺少 items 数组")
    meta = data.get("meta") or {}
    company = meta.get("company") or settings.BID_COMPANY_NAME
    profile = meta.get("company_profile") or {}
    # 原件 `file` 相对清单目录的 base_dir（默认同目录）解析为绝对路径。
    base_dir = (manifest_path.parent / str(meta.get("base_dir") or ".")).resolve()
    return {"company": str(company), "items": items,
            "profile": profile if isinstance(profile, dict) else {}, "base_dir": base_dir}


def _item_keywords(item: dict) -> list[str]:
    keywords = [str(k).strip() for k in (item.get("keywords") or []) if str(k).strip()]
    name = str(item.get("name") or "").strip()
    if name:
        keywords.append(name)
    # 归一化去重，保留顺序
    seen: set[str] = set()
    result = []
    for keyword in keywords:
        if keyword not in seen:
            seen.add(keyword)
            result.append(keyword)
    return result


def _needs_verification(item: dict) -> bool:
    for field in ("issuer", "cert_no", "valid_from", "valid_to"):
        if str(item.get(field) or "").strip() in ("", _UNVERIFIED):
            return True
    return False


def _match_requirement(text: str, items: list[dict]) -> list[tuple[dict, list[str]]]:
    hits: list[tuple[dict, list[str]]] = []
    for item in items:
        matched = [kw for kw in _item_keywords(item) if len(kw) >= 2 and kw in text]
        if matched:
            hits.append((item, matched))
    return hits


def _is_qualification_requirement(text: str) -> bool:
    """是否为“需要企业资质”的要求（证书/许可/认证等）。"""
    return bool(_QUALIFICATION_PATTERN.search(text))


def qualify_match(requirements: list[dict], manifest: dict) -> dict:
    """只对资质类要求做比对；其余（流程/承诺/报价等）标为“不适用”，不计入门槛缺口。"""
    items = manifest.get("items", [])
    results: list[dict] = []
    satisfied = unverified = missing = mandatory_missing = not_applicable = 0
    for requirement in requirements:
        text = str(requirement.get("text") or "").strip()
        category = str(requirement.get("category") or "其他")
        mandatory = bool(requirement.get("mandatory", True))
        if not _is_qualification_requirement(text):
            not_applicable += 1
            results.append({
                "requirement_id": str(requirement.get("id") or ""),
                "requirement_text": text,
                "category": category,
                "mandatory": mandatory,
                "status": STATUS_NA,
                "matched_qualification_ids": [],
                "matched_qualification_names": [],
                "reason": "非资质类要求（流程 / 承诺 / 商务等），不参与资质比对",
            })
            continue
        hits = _match_requirement(text, items) if text else []
        if not hits:
            status = STATUS_MISSING
            ids: list[str] = []
            names: list[str] = []
            reason = "未在企业资质清单中找到匹配项，疑似资质缺口，需人工确认"
            missing += 1
            if mandatory:
                mandatory_missing += 1
        else:
            ids = [str(item.get("id")) for item, _ in hits]
            names = [str(item.get("name")) for item, _ in hits]
            if any(_needs_verification(item) for item, _ in hits):
                status = STATUS_UNVERIFIED
                reason = f"命中资质 {'、'.join(names)}；证书编号 / 有效期等字段待人工核实"
                unverified += 1
            else:
                status = STATUS_SATISFIED
                reason = f"命中资质 {'、'.join(names)}"
                satisfied += 1
        results.append({
            "requirement_id": str(requirement.get("id") or ""),
            "requirement_text": text,
            "category": category,
            "mandatory": mandatory,
            "status": status,
            "matched_qualification_ids": ids,
            "matched_qualification_names": names,
            "reason": reason,
        })
    summary = {
        "total": len(results), "satisfied": satisfied, "unverified": unverified,
        "missing": missing, "mandatory_missing": mandatory_missing,
        "not_applicable": not_applicable,
    }
    return {"company": manifest.get("company", settings.BID_COMPANY_NAME),
            "results": results, "summary": summary}


# ---------- 定性可行性评估 ----------


def assess(match: dict) -> dict:
    """基于资质比对结论给出定性可行性评估（非概率）。"""
    summary = match.get("summary", {})
    mandatory_missing = int(summary.get("mandatory_missing", 0))
    missing = int(summary.get("missing", 0))
    unverified = int(summary.get("unverified", 0))
    total = int(summary.get("total", 0))

    reasons: list[str] = []
    risks: list[str] = []
    if total == 0:
        verdict = "信息不足"
        reasons.append("未提取到可比对的要求，无法判断资质匹配情况")
        risks.append("请确认招标文件正文已正确解析，或先在 Dify 中启用智能解读")
    elif mandatory_missing > 0:
        verdict = "不满足资格门槛"
        reasons.append(f"存在 {mandatory_missing} 项资格门槛要求未匹配到对应资质")
        risks.append("资格门槛缺口可能导致投标被否决，建议核实后再决定是否投标")
    elif missing > 0 or unverified > 0:
        verdict = "基本满足需补齐"
        reasons.append(f"存在 {missing} 项要求未匹配、{unverified} 项资质信息待核实")
        risks.append("请补齐缺失资质或核实资质有效期 / 编号后再投标")
    else:
        verdict = "具备投标可行性"
        reasons.append("全部要求均匹配到企业资质清单中的条目")
    if unverified:
        risks.append("资质清单中仍有字段标注为「待核实」，正式投标前须对照原件确认")

    return {"verdict": verdict, "reasons": reasons, "risks": risks, "disclaimer": _DISCLAIMER}


# ---------- 本地启发式解读（未配置 Dify 时兜底） ----------


def _categorize(line: str) -> str:
    if any(hint in line for hint in _QUALIFICATION_HINTS):
        return "资格"
    if any(hint in line for hint in ("技术", "参数", "规格", "功能", "方案")):
        return "技术"
    if any(hint in line for hint in ("报价", "商务", "付款", "交货", "服务", "工期")):
        return "商务"
    return "其他"


def local_extract(tender_text: str) -> dict:
    """启发式提取要点 / 要求 / 需插入资质位置。仅供本地联调，精度有限。"""
    lines = [line.strip() for line in re.split(r"[\r\n]+", tender_text or "")]
    lines = [line for line in lines if line]

    requirements: list[dict] = []
    key_points: list[dict] = []
    insertion_points: list[dict] = []
    for index, line in enumerate(lines, 1):
        if len(line) < 6:
            continue
        if any(hint in line for hint in _REQUIREMENT_HINTS):
            requirement_id = f"r{len(requirements) + 1}"
            requirements.append({
                "id": requirement_id,
                "text": line,
                "category": _categorize(line),
                "mandatory": any(hint in line for hint in _MANDATORY_HINTS),
            })
        if len(key_points) < 20 and (any(hint in line for hint in _QUALIFICATION_HINTS)
                                     or re.match(r"^[一二三四五六七八九十]+[、.．]|^\d+[、.．]", line)):
            key_points.append({"id": f"k{len(key_points) + 1}", "title": line[:24], "detail": line})
        if any(hint in line for hint in ("资格证明", "资质", "证明材料", "附件", "营业执照", "证书")):
            insertion_points.append({
                "id": f"p{len(insertion_points) + 1}",
                "location": line[:40],
                "expected_qualifications": [hint for hint in _QUALIFICATION_HINTS if hint in line][:3],
                "note": "疑似需插入企业资质的位置，请人工确认",
            })
    return {"key_points": key_points, "requirements": requirements, "insertion_points": insertion_points}


# ---------- Dify 工作流调用 ----------


def _dify_base() -> str:
    return settings.BID_DIFY_URL.strip().rstrip("/")


async def run_workflow(api_key: str, inputs: dict[str, Any]) -> dict:
    """调用 Dify Workflow 的 /workflows/run（blocking 模式），返回 outputs。"""
    base = _dify_base()
    if not base:
        raise BidError("未配置 BID_DIFY_URL")
    url = f"{base}/workflows/run"
    try:
        async with httpx.AsyncClient(timeout=settings.BID_DIFY_TIMEOUT, trust_env=False) as client:
            response = await client.post(
                url,
                headers={"Authorization": f"Bearer {api_key}"},
                json={"inputs": inputs, "response_mode": "blocking", "user": "bid-agent"},
            )
    except httpx.HTTPError as exc:
        raise BidError(f"Dify 调用失败：{type(exc).__name__}") from exc
    if response.status_code != 200:
        raise BidError(f"Dify 返回 HTTP {response.status_code}")
    try:
        payload = response.json()
    except ValueError as exc:
        raise BidError("Dify 返回非 JSON 响应") from exc
    data = payload.get("data") or {}
    if data.get("status") != "succeeded":
        raise BidError(f"Dify 工作流未成功：{data.get('error') or data.get('status')}")
    outputs = data.get("outputs") or {}
    if not isinstance(outputs, dict):
        raise BidError("Dify outputs 结构异常")
    return outputs


def _parse_json_output(outputs: dict, key: str = "result") -> dict:
    raw = outputs.get(key)
    if not isinstance(raw, str) or not raw.strip():
        raise BidError(f"Dify 输出缺少 {key}")
    try:
        parsed = json.loads(raw)
    except ValueError as exc:
        raise BidError(f"Dify 输出 {key} 不是合法 JSON") from exc
    if not isinstance(parsed, dict):
        raise BidError(f"Dify 输出 {key} 不是对象")
    return parsed


def _manifest_for_prompt(manifest: dict) -> str:
    compact = [{"id": item.get("id"), "name": item.get("name"), "category": item.get("category"),
                "keywords": item.get("keywords")} for item in manifest.get("items", [])]
    return json.dumps(compact, ensure_ascii=False)


async def analyze(tender_text: str, manifest_path: str | None = None) -> dict:
    """招标解读 + 资质比对 + 定性可行性评估。"""
    tender_text = tender_text or ""
    if not tender_text.strip():
        raise BidError("招标文件正文为空，请先解析招标文件")
    manifest = load_manifest(manifest_path)
    warnings: list[str] = []

    if _dify_base() and settings.BID_ANALYSIS_KEY:
        try:
            outputs = await run_workflow(settings.BID_ANALYSIS_KEY, {
                "tender_text": tender_text,
                "manifest": _manifest_for_prompt(manifest),
            })
            extracted = _parse_json_output(outputs)
            mode = "dify"
        except BidError as exc:
            logger.warning("Dify 解读失败，回退本地启发式：%s", exc)
            warnings.append(f"Dify 解读失败，已回退本地启发式：{exc}")
            extracted = local_extract(tender_text)
            mode = "local"
    else:
        extracted = local_extract(tender_text)
        mode = "local"
        warnings.append("未配置 Dify（BID_DIFY_URL / BID_ANALYSIS_KEY），当前为本地启发式解读，结果仅供联调。")

    requirements = extracted.get("requirements") or []
    match = qualify_match(requirements, manifest)
    assessment = assess(match)
    return {
        "mode": mode,
        "company": manifest["company"],
        "key_points": extracted.get("key_points") or [],
        "requirements": requirements,
        "insertion_points": extracted.get("insertion_points") or [],
        "match": {"success": True, "message": "比对完成", **match},
        "assessment": assessment,
        "warnings": warnings,
    }


def _seed_sections(company: str, profile: dict | None = None) -> list[dict]:
    """本地模式下的通用确定性内容（占位，须人工补充）。"""
    profile = profile or {}
    uscc = profile.get("unified_social_credit_code") or "（待补充）"
    legal = profile.get("legal_representative") or "（待补充）"
    return [
        {"title": "投标函", "volume": "business_technical",
         "content": f"致：（招标人）\n我方 {company} 参加本项目投标，愿按招标文件要求提供相关货物与服务。\n（投标函具体内容与签章由人工补充）"},
        {"title": "投标人基本情况", "volume": "business_technical",
         "content": f"投标人名称：{company}\n统一社会信用代码：{uscc}\n法定代表人：{legal}\n注册地址：（待补充）\n联系电话：（待补充）"},
    ]


def _resolve_attachments(attachments: list[dict], manifest: dict, missing: list[str],
                        base_dir: Path | None = None) -> list[dict]:
    items = {str(item.get("id")): item for item in manifest.get("items", [])}
    base = Path(base_dir or manifest.get("base_dir") or ".")
    resolved: list[dict] = []
    for attachment in attachments:
        volume = attachment.get("volume") or "business_technical"
        qualification_id = attachment.get("qualification_id")
        if qualification_id:
            item = items.get(str(qualification_id))
            if not item:
                missing.append(f"未找到资质 {qualification_id}")
                continue
            raw = Path(item.get("file") or "")
            path = _locate_attachment(base, raw)
            if path is None:
                missing.append(f"资质原件不存在：{item.get('name') or qualification_id}")
                continue
            resolved.append({"title": item.get("name"), "path": str(path), "volume": volume})
        else:
            raw = Path(attachment.get("path") or "")
            path = _locate_attachment(base, raw)
            if path is None:
                missing.append(f"附件不存在：{attachment.get('title') or raw.name}")
                continue
            resolved.append({"title": attachment.get("title") or path.stem, "path": str(path), "volume": volume})
    return resolved


async def draft(request: dict) -> dict:
    """根据招标文件与投标文件要求生成两册 Word 初稿。"""
    manifest = load_manifest(request.get("manifest_path"))
    company = manifest["company"]
    warnings: list[str] = []
    missing = list(request.get("missing") or [])

    sections = request.get("sections") or []
    mode = "local"
    if _dify_base() and settings.BID_DRAFT_KEY:
        try:
            outputs = await run_workflow(settings.BID_DRAFT_KEY, {
                "tender_text": request.get("tender_text") or "",
                "project_name": request.get("project_name") or "",
                "manifest": _manifest_for_prompt(manifest),
            })
            generated = _parse_json_output(outputs)
            if generated.get("sections"):
                sections = generated["sections"]
                mode = "dify"
        except BidError as exc:
            logger.warning("Dify 初稿生成失败，回退本地章节：%s", exc)
            warnings.append(f"Dify 初稿生成失败，已回退本地章节：{exc}")

    if not sections:
        sections = _seed_sections(company, manifest.get("profile"))
        warnings.append("未提供章节内容且未启用 Dify，已生成通用占位章节，须人工补充。")

    attachments = _resolve_attachments(request.get("attachments") or [], manifest, missing, manifest.get("base_dir"))
    project = {"name": request.get("project_name") or "投标文件", "bid_no": request.get("bid_no")}

    output_dir = Path(settings.BID_OUTPUT_DIR)
    result = build_bid_documents(output_dir, project, sections, attachments, None,
                                 final=bool(request.get("final")), missing=missing)
    files = [{"name": item["name"], "kind": item["kind"], "download_url": f"/api/v1/bid/files/{item['name']}"}
             for item in result["files"]]
    return {"mode": mode, "files": files, "warnings": result["warnings"] + warnings}


def _ocr_command() -> str | list[str]:
    """OCR 命令：.py 适配脚本用当前解释器执行，其它按可执行文件处理。"""
    command = (settings.BID_OCR_COMMAND or "").strip()
    if not command:
        return ""
    script = Path(command)
    if script.suffix.lower() == ".py":
        return [sys.executable, str(script)]
    return command


def parse_upload(path: Path) -> dict:
    """解析上传的原件，返回 blocks / text / warnings / complete。

    配置 BID_DOC_SERVICE_URL 时，.doc 转换与扫描 PDF OCR 走 doc-service；否则回退本地二进制。
    """
    result = parse_document(path, converter=settings.BID_LIBREOFFICE, ocr_command=_ocr_command(),
                            converter_remote=settings.BID_DOC_SERVICE_URL,
                            ocr_remote=settings.BID_DOC_SERVICE_URL,
                            remote_token=settings.BID_DOC_SERVICE_TOKEN)
    blocks = result.get("blocks", [])
    text = "\n".join(block["text"] for block in blocks)
    return {"blocks": blocks, "text": text, "char_count": len(text),
            "warnings": result.get("warnings", []), "complete": bool(result.get("complete", True))}


__all__ = ["BidError", "DocumentError", "load_manifest", "qualify_match", "assess",
           "local_extract", "analyze", "draft", "parse_upload", "run_workflow"]
