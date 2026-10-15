"""废标项 / 无效标检查服务：程序维护权威状态，模型只提交增量 patch。

设计要点（针对长投标文件）：
- 检查按「文档 × 片段」顺序推进，每个片段只提交增量：证据、新增待定风险、
  更新、解除。由本模块合并进权威状态，模型从不重写整份状态。
- **缺失类结论只能进「待定」**：只要结论是「投标文件未提供 / 未响应」，或没有
  给出可引用的投标证据，程序一律降级为待定，等后续片段推翻或补证。
- **结构性线索防误判**：证据/理由中出现「目录 / 章节标题 / 附件 / 清单 / 页码」
  等线索时，不得因「看不到正文」判定缺失，同样降级为待定。
- 待定风险在补到证据后可**升级**为已确认风险；补不到则可在收尾时保持待定。

状态以 JSON 落盘（`BID_OUTPUT_DIR/rejection/<session_id>.json`），进程重启后可续跑。
本模块不调用任何模型：所有判定输入都来自调用方提交的 patch。
"""
from __future__ import annotations

import io
import json
import os
import re
import tempfile
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .settings import settings


class RejectionError(ValueError):
    """废标项检查模块的可预期错误（对应 HTTP 422）。"""


STATE_VERSION = 1

# 结论类型：无效投标 / 废标项
FINDING_TYPES = ("invalidBid", "rejectionItem")
FINDING_TYPE_LABELS = {"invalidBid": "无效投标", "rejectionItem": "废标项"}
# 严重度
SEVERITIES = ("high", "medium", "low")
SEVERITY_LABELS = {"high": "高", "medium": "中", "low": "低"}
SEVERITY_ORDER = {"high": 0, "medium": 1, "low": 2}
# 结论口径：present=已发现的问题内容；absent=判定“未提供/未响应”
CLAIMS = ("present", "absent")

_MISSING_WORDS = ("缺失", "未见", "未提供", "未响应", "未包含", "没有", "无对应", "未发现", "遗漏", "漏项")

# 纸质 / 线下事项：电子文件无法判定，默认不纳入检查（否则必然误报）。
# 硬排除：这些要求无法由电子文件满足，出现在标题或依据里都算线下事项。
_HARD_OFFLINE_PATTERN = re.compile(
    r"密封|纸质|现场递交|现场提交|现场核验|装订|骑缝|光盘|刻录|U\s?盘|优盘"
)
# 软排除：签字/盖章是纸质动作，但常出现在「格式须加盖公章」这类仍可电子核查的
# 要求里，因此只在**标题**（即判定动作本身）出现时才判为线下事项。
_SOFT_OFFLINE_PATTERN = re.compile(r"签字|盖章|签章|加盖|签属|手写")

# 结构性线索：出现这些说明文档结构确实存在，不得据此判定“缺少该章节/附件”。
_STRUCTURAL_PATTERN = re.compile(
    r"目录|章节|标题|附件|附表|清单|页码|第\s*[一二三四五六七八九十百零〇\d]+\s*[章节篇部分条]"
)

_SESSION_ID = re.compile(r"^rej-\d{8}-[a-f0-9]{8}$")


# ---------- 通用小工具 ----------


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _new_id(prefix: str, number: int) -> str:
    return f"{prefix}_{number:04d}"


def _next_id(state: dict, key: str, prefix: str) -> str:
    counters = state.setdefault("counters", {})
    counters[key] = int(counters.get(key, 0)) + 1
    return _new_id(prefix, counters[key])


def _clip(value: Any, limit: int) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    return text[:limit]


def has_structural_clue(text: str) -> bool:
    """文本是否含目录/章节/附件/清单/页码等结构性线索。"""
    return bool(_STRUCTURAL_PATTERN.search(text or ""))


def _is_missing_claim(text: str) -> bool:
    return any(word in (text or "") for word in _MISSING_WORDS)


def classify_scope(title: str, requirement: str = "") -> str:
    """判断一条要求属于电子文件可检查（content）还是纸质/线下事项（offline）。

    硬性线下标记（密封/纸质/现场递交/装订…）出现在标题或依据里都判线下；
    签字/盖章类软标记只在标题出现时判线下，避免把「格式须加盖公章」这类
    仍可电子核查的格式要求误排除。
    """
    if _HARD_OFFLINE_PATTERN.search(f"{title}{requirement}"):
        return "offline"
    if _SOFT_OFFLINE_PATTERN.search(title or ""):
        return "offline"
    return "content"


# ---------- 会话状态存取 ----------


def session_dir() -> Path:
    """检查会话目录（与 Word 输出同卷，容器里同样可写）。"""
    return Path(settings.bid_state_dir) / "rejection"


def _state_path(session_id: str) -> Path:
    if not _SESSION_ID.fullmatch(str(session_id or "")):
        raise RejectionError(f"session_id 非法：{session_id!r}")
    return session_dir() / f"{session_id}.json"


def save_state(state: dict) -> Path:
    """原子写入会话状态（临时文件 + rename）。"""
    path = _state_path(str(state["session_id"]))
    path.parent.mkdir(parents=True, exist_ok=True)
    state["updated_at"] = _now()
    payload = json.dumps(state, ensure_ascii=False, indent=2)
    handle_fd, temporary = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    try:
        with os.fdopen(handle_fd, "w", encoding="utf-8") as handle:
            handle.write(payload)
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)
    return path


def load_state(session_id: str) -> dict:
    path = _state_path(session_id)
    if not path.is_file():
        raise RejectionError(f"未找到检查会话：{session_id}")
    try:
        state = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise RejectionError(f"检查会话无法解析：{type(exc).__name__}") from exc
    if not isinstance(state, dict) or int(state.get("version") or 0) != STATE_VERSION:
        raise RejectionError("检查会话版本不匹配，请重新初始化")
    return state


# ---------- 初始化 ----------


def normalize_checklist(items: list[dict]) -> list[dict]:
    """归一化模型从招标文件抽取的无效投标 / 废标项清单：稳定 id、去重、标注线下事项。"""
    if not isinstance(items, list):
        raise RejectionError("rejection_items 必须是数组")
    checklist: list[dict] = []
    seen: set[str] = set()
    for raw in items:
        if not isinstance(raw, dict):
            continue
        title = _clip(raw.get("title") or raw.get("name"), 80)
        requirement = _clip(raw.get("requirement") or raw.get("text") or raw.get("detail") or title, 600)
        if not title and not requirement:
            continue
        key = re.sub(r"\s+", "", title or requirement)[:60]
        if key in seen:
            continue
        seen.add(key)

        finding_type = str(raw.get("type") or "").strip()
        if finding_type not in FINDING_TYPES:
            blob = f"{title}{requirement}"
            finding_type = "invalidBid" if ("无效" in blob or "否决" in blob) else "rejectionItem"

        scope = classify_scope(title, requirement)
        checklist.append({
            "id": _new_id("ri", len(checklist) + 1),
            "type": finding_type,
            "title": title or requirement[:80],
            "requirement": requirement,
            "source": _clip(raw.get("source"), 120),
            "scope": scope,
            "applicable": scope == "content",
        })
    if not checklist:
        raise RejectionError("废标项清单为空：请先提交从招标文件抽取的无效投标 / 废标项")
    return checklist


def normalize_documents(documents: list[dict] | None) -> list[dict]:
    """归一化待检查的投标文件清单。"""
    if documents is None:
        return []
    if not isinstance(documents, list):
        raise RejectionError("documents 必须是数组")
    normalized: list[dict] = []
    for raw in documents:
        if isinstance(raw, str):
            raw = {"name": raw}
        if not isinstance(raw, dict):
            continue
        name = _clip(raw.get("name") or raw.get("title"), 120)
        if not name:
            continue
        normalized.append({
            "id": _new_id("doc", len(normalized) + 1),
            "name": name,
            "chunk_total": max(1, int(raw.get("chunk_total") or 1)),
            "chunks_done": 0,
        })
    return normalized


def init_session(project: dict | None, rejection_items: list[dict],
                 documents: list[dict] | None = None) -> dict:
    """创建一次废标项检查会话，返回初始状态。"""
    project = project if isinstance(project, dict) else {}
    session_id = f"rej-{time.strftime('%Y%m%d')}-{uuid.uuid4().hex[:8]}"
    state = {
        "version": STATE_VERSION,
        "session_id": session_id,
        "project": {"name": _clip(project.get("name") or project.get("project_name"), 120),
                    "bid_no": _clip(project.get("bid_no"), 60)},
        "created_at": _now(),
        "updated_at": _now(),
        "checklist": normalize_checklist(rejection_items),
        "documents": normalize_documents(documents),
        "cursor": {"document_id": None, "chunk_index": 0, "finished": False},
        "evidence": [],
        "pending_risks": [],
        "resolved_risks": [],
        "confirmed_risks": [],
        "rounds": [],
        "counters": {"ev": 0, "risk": 0},
    }
    save_state(state)
    return state


# ---------- 风险条目归一化与判定 ----------


def _normalize_risk(state: dict, payload: dict, round_no: int) -> tuple[dict | None, str]:
    """把 patch 里的风险条目归一化为内部结构；返回 (risk, 拒绝原因)。"""
    if not isinstance(payload, dict):
        return None, "风险条目不是对象"
    checklist_ids = {item["id"] for item in state["checklist"]}
    checklist_id = _clip(payload.get("checklist_id"), 40)
    if checklist_id and checklist_id not in checklist_ids:
        return None, f"checklist_id 不存在：{checklist_id}"

    title = _clip(payload.get("title"), 80)
    requirement = _clip(payload.get("requirement"), 600)
    if not title or not requirement:
        return None, "风险条目缺少 title 或 requirement（检查依据）"

    evidence = _clip(payload.get("bid_evidence"), 600)
    finding_type = str(payload.get("type") or "").strip()
    if finding_type not in FINDING_TYPES:
        finding_type = "rejectionItem"
    severity = str(payload.get("severity") or "").strip()
    if severity not in SEVERITIES:
        severity = "high"

    claim = str(payload.get("claim") or "").strip()
    if claim not in CLAIMS:
        claim = "absent" if (not evidence or _is_missing_claim(f"{title}{evidence}")) else "present"

    return {
        "id": "",
        "checklist_id": checklist_id,
        "type": finding_type,
        "severity": severity,
        "claim": claim,
        "title": title,
        "requirement": requirement,
        "bid_evidence": evidence,
        "risk_reason": _clip(payload.get("risk_reason"), 600),
        "suggestion": _clip(payload.get("suggestion"), 400),
        "created_round": round_no,
        "history": [],
    }, ""


def _confirmable(risk: dict) -> tuple[bool, str]:
    """已确认风险必须自带可引用的投标证据，且不得是「缺失/未响应」类结论。

    「缺失」之所以不能直接确认：文件后续片段可能包含对应内容，程序无法证明
    「全文都没有」。这类结论一律留在待定，由后续片段补证升级或人工核对。
    """
    evidence = str(risk.get("bid_evidence") or "").strip()
    if risk.get("claim") == "absent":
        if has_structural_clue(f"{evidence}{risk.get('risk_reason')}"):
            return False, "证据含目录/章节/附件/清单/页码等结构性线索，不得据「看不到正文」判定缺失"
        return False, "缺失/未响应类结论不能直接确认，须先补到可引用的投标证据"
    if not evidence:
        return False, "缺少投标证据（bid_evidence），不能确认为风险"
    return True, ""


# ---------- 应用增量 patch ----------


def apply_patch(state: dict, patch: dict) -> dict:
    """把单个片段的增量合并进状态，返回 (状态, 诊断) —— 状态已就地更新。"""
    if not isinstance(patch, dict):
        raise RejectionError("patch 必须是对象")

    round_no = len(state.get("rounds", [])) + 1
    applied = {"evidence": 0, "pending_added": 0, "updated": 0, "resolved": 0,
               "confirmed_added": 0, "promoted": 0, "demoted": 0}
    rejected: list[dict] = []

    pending_by_id = {risk["id"]: risk for risk in state["pending_risks"]}

    # 1) 证据
    for entry in patch.get("evidence_adds") or []:
        if not isinstance(entry, dict):
            rejected.append({"kind": "evidence_add", "reason": "证据条目不是对象"})
            continue
        quote = _clip(entry.get("quote"), 600)
        locator = _clip(entry.get("locator"), 80)
        if not quote:
            rejected.append({"kind": "evidence_add", "reason": "证据缺少 quote（原文摘录）"})
            continue
        state["evidence"].append({
            "id": _next_id(state, "ev", "ev"),
            "checklist_id": _clip(entry.get("checklist_id"), 40),
            "document_id": _clip(entry.get("document_id") or patch.get("document_id"), 40),
            "locator": locator,
            "quote": quote,
            "round": round_no,
        })
        applied["evidence"] += 1

    # 2) 新增待定风险
    for entry in patch.get("pending_risk_adds") or []:
        risk, reason = _normalize_risk(state, entry if isinstance(entry, dict) else {}, round_no)
        if risk is None:
            rejected.append({"kind": "pending_risk_add", "reason": reason})
            continue
        risk["id"] = _next_id(state, "risk", "risk")
        risk["status"] = "pending"
        risk["pending_reason"] = "片段判定为疑似风险，待后续片段确认或排除"
        state["pending_risks"].append(risk)
        pending_by_id[risk["id"]] = risk
        applied["pending_added"] += 1

    # 3) 更新待定风险（可借此补证据 → 自动升级）
    for entry in patch.get("pending_risk_updates") or []:
        if not isinstance(entry, dict):
            rejected.append({"kind": "pending_risk_update", "reason": "更新条目不是对象"})
            continue
        target = pending_by_id.get(_clip(entry.get("id"), 40))
        if target is None:
            rejected.append({"kind": "pending_risk_update",
                             "reason": f"待定风险不存在：{entry.get('id')!r}"})
            continue
        snapshot = {key: target[key] for key in ("severity", "claim", "bid_evidence",
                                                 "risk_reason", "suggestion")}
        for field in ("severity", "claim", "bid_evidence", "risk_reason", "suggestion", "title", "requirement"):
            if field in entry:
                limit = {"bid_evidence": 600, "risk_reason": 600, "suggestion": 400,
                         "title": 80, "requirement": 600}.get(field, 60)
                target[field] = _clip(entry.get(field), limit)
        if target.get("severity") not in SEVERITIES:
            target["severity"] = "high"
        if target.get("claim") not in CLAIMS:
            target["claim"] = "present" if target.get("bid_evidence") else "absent"
        target["history"].append({"round": round_no, "before": snapshot})
        applied["updated"] += 1

        ok, reason = _confirmable(target)
        if ok:
            state["pending_risks"].remove(target)
            pending_by_id.pop(target["id"], None)
            target["status"] = "confirmed"
            target["confirmed_round"] = round_no
            target.pop("pending_reason", None)
            state["confirmed_risks"].append(target)
            applied["promoted"] += 1
        else:
            target["pending_reason"] = reason

    # 4) 解除待定风险（在别处找到了对应内容 / 判定与招标无关）
    for entry in patch.get("pending_risk_resolves") or []:
        if not isinstance(entry, dict):
            rejected.append({"kind": "pending_risk_resolve", "reason": "解除条目不是对象"})
            continue
        target = pending_by_id.get(_clip(entry.get("id"), 40))
        if target is None:
            rejected.append({"kind": "pending_risk_resolve",
                             "reason": f"待定风险不存在：{entry.get('id')!r}"})
            continue
        state["pending_risks"].remove(target)
        pending_by_id.pop(target["id"], None)
        target["status"] = "resolved"
        target["resolution"] = _clip(entry.get("resolution"), 400) or "后续片段已排除该疑点"
        target["resolved_round"] = round_no
        target.pop("pending_reason", None)
        state["resolved_risks"].append(target)
        applied["resolved"] += 1

    # 5) 直接确认风险（不满足可确认条件的一律降级为待定）
    for entry in patch.get("confirmed_risk_adds") or []:
        risk, reason = _normalize_risk(state, entry if isinstance(entry, dict) else {}, round_no)
        if risk is None:
            rejected.append({"kind": "confirmed_risk_add", "reason": reason})
            continue
        risk["id"] = _next_id(state, "risk", "risk")
        ok, why = _confirmable(risk)
        if ok:
            risk["status"] = "confirmed"
            risk["confirmed_round"] = round_no
            state["confirmed_risks"].append(risk)
            applied["confirmed_added"] += 1
        else:
            risk["status"] = "pending"
            risk["pending_reason"] = why
            state["pending_risks"].append(risk)
            pending_by_id[risk["id"]] = risk
            applied["demoted"] += 1

    # 6) 推进游标
    cursor = state["cursor"]
    if patch.get("document_id"):
        cursor["document_id"] = _clip(patch.get("document_id"), 40)
    if isinstance(patch.get("chunk_index"), int):
        cursor["chunk_index"] = patch["chunk_index"]
    if "finished" in patch:
        cursor["finished"] = bool(patch.get("finished"))
    if cursor.get("document_id"):
        for document in state["documents"]:
            if document["id"] == cursor["document_id"] and isinstance(patch.get("chunk_index"), int):
                document["chunks_done"] = max(document["chunks_done"], int(patch["chunk_index"]) + 1)

    state.setdefault("rounds", []).append({
        "round": round_no,
        "document_id": cursor.get("document_id"),
        "chunk_index": cursor.get("chunk_index"),
        "locator_range": _clip(patch.get("locator_range"), 80),
        "applied": applied,
        "rejected": rejected,
    })

    diagnostics = {"round": round_no, "applied": applied, "rejected": rejected,
                   "stats": state_stats(state), "cursor": dict(cursor)}
    return diagnostics


def state_stats(state: dict) -> dict:
    by_severity = {"high": 0, "medium": 0, "low": 0}
    for risk in state.get("confirmed_risks", []):
        by_severity[risk.get("severity", "high")] = by_severity.get(risk.get("severity", "high"), 0) + 1
    checklist = state.get("checklist", [])
    documents = state.get("documents", [])
    return {
        "checklist_total": len(checklist),
        "applicable": sum(1 for item in checklist if item.get("applicable")),
        "offline": sum(1 for item in checklist if not item.get("applicable")),
        "evidence": len(state.get("evidence", [])),
        "confirmed": len(state.get("confirmed_risks", [])),
        "pending": len(state.get("pending_risks", [])),
        "resolved": len(state.get("resolved_risks", [])),
        "by_severity": by_severity,
        "coverage": {
            "documents": len(documents),
            "chunks_total": sum(int(document.get("chunk_total") or 0) for document in documents),
            "chunks_done": sum(int(document.get("chunks_done") or 0) for document in documents),
            "finished": bool(state.get("cursor", {}).get("finished")),
            "rounds": len(state.get("rounds", [])),
        },
    }


# ---------- 收尾 ----------


def _dedupe(risks: list[dict]) -> list[dict]:
    seen: set[str] = set()
    result: list[dict] = []
    for risk in risks:
        compact_title = re.sub(r"\s+", "", str(risk.get("title") or ""))
        key = "{}|{}".format(risk.get("checklist_id"), compact_title)
        if key in seen:
            continue
        seen.add(key)
        result.append(risk)
    return result


def _sort_risks(risks: list[dict]) -> list[dict]:
    return sorted(risks, key=lambda risk: (SEVERITY_ORDER.get(risk.get("severity", "high"), 9),
                                           risk.get("id", "")))


def finalize(state: dict) -> dict:
    """收尾：合并去重、按严重度排序，产出报告（并写回状态）。"""
    confirmed = _sort_risks(_dedupe(state.get("confirmed_risks", [])))
    pending = _sort_risks(_dedupe(state.get("pending_risks", [])))
    resolved = _sort_risks(state.get("resolved_risks", []))
    warning = []
    if not state.get("cursor", {}).get("finished"):
        warning.append("检查尚未标记完成（cursor.finished=false），结论可能只覆盖了部分片段。")
    if pending:
        warning.append(f"仍有 {len(pending)} 条待定风险未决：多为「缺失/未响应」类结论，需人工核对原文或补证。")
    if any(not item.get("applicable") for item in state.get("checklist", [])):
        warning.append("清单含纸质/线下事项，已排除出检查范围，请人工按线下流程核对。")

    state["confirmed_risks"] = confirmed
    state["pending_risks"] = pending
    state["report"] = {
        "session_id": state["session_id"],
        "project": state.get("project", {}),
        "generated_at": _now(),
        "summary": state_stats(state),
        "findings": confirmed,
        "pending": pending,
        "resolved": resolved,
        "checklist": state.get("checklist", []),
        "warnings": warning,
    }
    save_state(state)
    return state["report"]


# ---------- 导出 xlsx ----------

_HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
_HEADER_FONT = Font(bold=True)
_WRAP = Alignment(wrap_text=True, vertical="top")

FINDING_COLUMNS = [
    ("序号", 6), ("类型", 10), ("严重度", 8), ("检查项", 30), ("检查依据", 46),
    ("投标证据", 46), ("风险原因", 34), ("处理建议", 30), ("关联检查项", 12), ("风险ID", 12),
]


def _write_sheet(sheet, columns: list[tuple[str, int]], rows: list[list[Any]]) -> None:
    sheet.append([title for title, _ in columns])
    for cell, (_, width) in zip(sheet[1], columns):
        cell.fill, cell.font = _HEADER_FILL, _HEADER_FONT
        sheet.column_dimensions[cell.column_letter].width = width
    for row in rows:
        sheet.append(row)
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = _WRAP
    sheet.freeze_panes = "A2"


def _finding_row(index: int, risk: dict) -> list[Any]:
    return [
        index,
        FINDING_TYPE_LABELS.get(risk.get("type"), risk.get("type") or ""),
        SEVERITY_LABELS.get(risk.get("severity"), risk.get("severity") or ""),
        risk.get("title") or "",
        risk.get("requirement") or "",
        risk.get("bid_evidence") or "",
        risk.get("risk_reason") or "",
        risk.get("suggestion") or "",
        risk.get("checklist_id") or "",
        risk.get("id") or "",
    ]


def build_workbook(report: dict, extra_sheets: list[dict] | None = None,
                   extra_overview_rows: list[tuple] | None = None) -> bytes:
    """把报告写成 xlsx 字节流。

    `extra_sheets` 与 `extra_overview_rows` 供错别字 / 逻辑谬误检查复用**同一出口**
    （它们与废标项检查共用这份报告）。
    """
    if not isinstance(report, dict):
        raise RejectionError("报告结构异常")
    summary = report.get("summary") or {}
    cover = summary.get("coverage") or {}
    workbook = Workbook()

    overview = workbook.active
    overview.title = "检查概览"
    overview.append(["指标", "值"])
    for cell in overview[1]:
        cell.fill, cell.font = _HEADER_FILL, _HEADER_FONT
    overview.column_dimensions["A"].width, overview.column_dimensions["B"].width = 24, 46
    rows = [
        ("项目名称", (report.get("project") or {}).get("name") or ""),
        ("检查会话", report.get("session_id") or ""),
        ("生成时间", report.get("generated_at") or ""),
        ("检查项总数", summary.get("checklist_total", 0)),
        ("其中可电子检查", summary.get("applicable", 0)),
        ("其中纸质/线下（已排除）", summary.get("offline", 0)),
        ("已确认风险", summary.get("confirmed", 0)),
        ("高风险", (summary.get("by_severity") or {}).get("high", 0)),
        ("中风险", (summary.get("by_severity") or {}).get("medium", 0)),
        ("低风险", (summary.get("by_severity") or {}).get("low", 0)),
        ("待定风险", summary.get("pending", 0)),
        ("已排除疑点", summary.get("resolved", 0)),
        ("证据条数", summary.get("evidence", 0)),
        ("投标文件数", cover.get("documents", 0)),
        ("片段进度", f"{cover.get('chunks_done', 0)}/{cover.get('chunks_total', 0)}"),
        ("检查已完结", "是" if cover.get("finished") else "否"),
    ]
    rows.extend(extra_overview_rows or [])
    for key, value in rows:
        overview.append([key, value])
    for note in report.get("warnings") or []:
        overview.append(["提示", note])

    findings = report.get("findings") or []
    _write_sheet(workbook.create_sheet("废标项"), FINDING_COLUMNS,
                 [_finding_row(index, risk) for index, risk in enumerate(findings, 1)])

    pending = report.get("pending") or []
    pending_columns = FINDING_COLUMNS[:-2] + [("待定原因", 40), ("风险ID", 12)]
    _write_sheet(workbook.create_sheet("待定风险"), pending_columns, [
        _finding_row(index, risk)[:-2] + [risk.get("pending_reason") or "", risk.get("id") or ""]
        for index, risk in enumerate(pending, 1)
    ])

    for sheet_spec in extra_sheets or []:
        _write_sheet(workbook.create_sheet(_clip(sheet_spec.get("title"), 28) or "附加"),
                     sheet_spec.get("columns") or FINDING_COLUMNS,
                     sheet_spec.get("rows") or [])

    stream = io.BytesIO()
    workbook.save(stream)
    return stream.getvalue()


def export_workbook(report: dict, *, extra_sheets: list[dict] | None = None,
                   extra_overview_rows: list[tuple] | None = None, name: str = "检查报告.xlsx") -> dict:
    """通用导出：把报告（可含错别字 / 逻辑谬误的附加表）写成 xlsx。"""
    directory = Path(settings.BID_OUTPUT_DIR)
    directory.mkdir(parents=True, exist_ok=True)
    safe_name = Path(str(name or "检查报告.xlsx")).name
    if not safe_name.lower().endswith(".xlsx"):
        safe_name += ".xlsx"
    path = directory / safe_name
    handle_fd, temporary = tempfile.mkstemp(dir=directory, suffix=".xlsx")
    try:
        os.close(handle_fd)
        Path(temporary).write_bytes(build_workbook(report, extra_sheets, extra_overview_rows))
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)
    return {"name": safe_name, "path": str(path.resolve())}


def export_report(session_id: str, report: dict | None = None) -> dict:
    """导出废标项检查报告为 xlsx，返回文件名与路径（调用方负责生成下载链接）。"""
    if report is None:
        state = load_state(session_id)
        report = state.get("report") or finalize(state)
    return export_workbook(report, name=f"{session_id}-废标项检查.xlsx")


__all__ = ["RejectionError", "FINDING_TYPES", "FINDING_TYPE_LABELS", "SEVERITIES", "SEVERITY_LABELS",
           "classify_scope", "has_structural_clue", "init_session", "load_state", "save_state",
           "apply_patch", "finalize", "build_workbook", "export_workbook", "export_report",
           "session_dir", "state_stats"]
