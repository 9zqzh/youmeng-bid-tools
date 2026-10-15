"""正文检查（错别字 / 逻辑谬误）：模型提线索，**程序回验**，与废标项检查共用同一导出出口。

为什么程序必须回验：

- **错别字会「造」出来**。模型对专业术语很容易误判，把正确写法当成错别字。所以每个候选项都
  回到原文里查证：**定位不到的一律不算数**（单独列进 `unlocated` 而不是当真）。另外若「错别字」
  文本本身出现在**招标原文**里，那大概率是招标方的既有用词 → 降级为「待核对」，
  而不是判成错误（与查重的招标白名单同一个道理，这里直接复用它的招标索引）。
- **逻辑谬误靠事实登记表累积**。逐段上报的类型化事实按 (分类, 名称) 归集，跨段取值不一致即判
  前后矛盾。这部分直接复用一致性审计的台账实现，避免两套逻辑各自演化。

导出：`sheets_for()` 产出的附加表可直接喂给 `bid_rejection.build_workbook`，
于是错别字 / 逻辑谬误 / 废标项检查最终落在**同一份 xlsx** 里。
"""
from __future__ import annotations

import json
import os
import re
import tempfile
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

from .settings import settings
from . import bid_consistency, bid_duplicate, bid_rejection


class ReviewError(ValueError):
    """正文检查模块的可预期错误（对应 HTTP 422）。"""


STATE_VERSION = 1
REVIEW_ID = re.compile(r"^rev-\d{8}-[a-f0-9]{8}$")

# 错别字候选的四种归宿
TYPO_STATUS = ("confirmed", "needs_review", "unlocated", "invalid")
TYPO_STATUS_LABELS = {"confirmed": "确认（原文已定位）", "needs_review": "待核对（招标既有用词）",
                      "unlocated": "未定位（疑似误判）", "invalid": "无效（参数不成立）"}

MAX_CANDIDATES = 500
MAX_LOCATIONS = 20

TYPO_COLUMNS = [("序号", 6), ("位置(行:列)", 14), ("所在行", 46), ("错别字", 16), ("建议更正", 16),
                ("说明", 30), ("状态", 22), ("出现次数", 10), ("条目ID", 12)]
LOGIC_COLUMNS = [("序号", 6), ("类型", 12), ("分类", 14), ("事实/要点", 30), ("涉及小节", 34),
                 ("冲突取值", 34), ("说明", 50), ("条目ID", 12)]


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _clip(value: Any, limit: int) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()[:limit]


# ---------- 通用：原文定位 ----------


def locate_text(source: str, needle: str, *, limit: int = MAX_LOCATIONS) -> list[dict]:
    """在原文里逐个定位 needle，返回 1 基的行号 / 列号 / 所在行。"""
    text = str(source or "")
    target = str(needle or "")
    if not target:
        return []
    hits: list[dict] = []
    start = 0
    while len(hits) < limit:
        index = text.find(target, start)
        if index < 0:
            break
        line_number = text.count("\n", 0, index) + 1
        line_start = text.rfind("\n", 0, index) + 1
        line_end = text.find("\n", index)
        line_text = text[line_start:line_end if line_end >= 0 else len(text)]
        hits.append({"line": line_number, "column": index - line_start + 1,
                     "offset": index, "line_text": _clip(line_text, 160)})
        start = index + max(1, len(target))
    return hits


# ---------- 错别字 ----------


def _normalize_candidates(raw: Any) -> list[dict]:
    if raw is None:
        return []
    if isinstance(raw, dict):
        raw = raw.get("candidates") if isinstance(raw.get("candidates"), list) else [raw]
    if not isinstance(raw, list):
        raise ReviewError("candidates 必须是数组")
    if len(raw) > MAX_CANDIDATES:
        raise ReviewError(f"单次最多回验 {MAX_CANDIDATES} 条候选（当前 {len(raw)}）")
    normalized: list[dict] = []
    for index, item in enumerate(raw, 1):
        if isinstance(item, str):
            item = {"wrong": item}
        if not isinstance(item, dict):
            continue
        normalized.append({
            "id": _clip(item.get("id"), 40) or f"t{index:04d}",
            "wrong": _clip(item.get("wrong") or item.get("text"), 80),
            "right": _clip(item.get("right") or item.get("suggestion"), 80),
            "context": _clip(item.get("context") or item.get("quote"), 300),
            "note": _clip(item.get("note") or item.get("reason"), 200),
        })
    return normalized


def _tender_verdict(text: str, normalized_tender: str, skeleton_tender: str) -> str:
    """判断一段文字是否「就是招标原文里的用词」。

    这里用**包含**而不是句子级相似度：错别字候选往往是招标原文某个句子里的**片段**
    （例如「编制投标文件」），整句比对根本命中不了。所以直接拿归一化后的招标全文做子串判断。
    只差数字的情况（如「投标有效期 90 天」vs「120 天」）用骨架再判一次 —— 用词仍是招标的，
    但改的是数字，属于内容差异而非错别字，所以单独记为 skeleton 供调用方区分。
    """
    if not normalized_tender:
        return "none"
    normalized = bid_duplicate.normalize_sentence(text)
    if not normalized:
        return "none"
    if normalized in normalized_tender:
        return "exact"
    if bid_duplicate.skeleton_sentence(normalized) in skeleton_tender:
        return "skeleton"
    return "none"


def check_typos(source_text: Any = None, candidates: Any = None, tender_text: Any = None,
                source_name: str = "", source_path: Any = None, tender_path: Any = None) -> dict:
    """回验错别字候选：必须能在原文定位；招标既有用词降级为待核对。

    正文与招标原文都可**给文本或给服务器上的文件路径**——大文档走路径更稳
    （避免把整篇正文塞进工具参数而被截断）。
    """
    resolved_name = _clip(source_name, 120)
    if source_text is None and source_path:
        source_text = bid_duplicate.read_document_text(source_path)
        resolved_name = resolved_name or Path(str(source_path)).name
    if tender_text is None and tender_path:
        tender_text = bid_duplicate.read_document_text(tender_path)

    source = str(source_text or "")
    if not source.strip():
        raise ReviewError("source_text / source_path 至少要给一个，用于回验错别字")
    items = _normalize_candidates(candidates)
    if not items:
        raise ReviewError("candidates 为空：请先提交模型给出的错别字候选")

    normalized_tender = bid_duplicate.normalize_sentence(tender_text) if tender_text else ""
    skeleton_tender = bid_duplicate.skeleton_sentence(tender_text) if tender_text else ""

    findings: list[dict] = []
    unlocated: list[dict] = []
    invalid: list[dict] = []
    for item in items:
        wrong, right = item["wrong"], item["right"]
        if not wrong or not right or wrong == right:
            invalid.append({**item, "status": "invalid",
                            "reason": "缺少错别字或建议值，或两者相同"})
            continue
        hits = locate_text(source, wrong)
        if not hits:
            # 关键：定位不到就不当数（防模型造错）
            unlocated.append({**item, "status": "unlocated", "occurrences": 0,
                              "reason": f"原文中找不到「{wrong}」，疑似误判，已排除"})
            continue

        # 定位：优先取「落在所给上下文之内」的那一次出现（上下文可用于消歧），
        # 否则回退到第一次出现。定位点始终指向**错别字本身**，不是上下文开头。
        context_hits = locate_text(source, item["context"]) if item["context"] else []
        occurrence = hits[0]
        context_verified = False
        context_note = ""
        if item["context"]:
            if not context_hits:
                context_note = "给出的上下文未能在原文定位到，位置按错别字本身推算"
            else:
                span_start = context_hits[0]["offset"]
                span_end = span_start + len(item["context"])
                inside = [hit for hit in hits if span_start <= hit["offset"] < span_end]
                if inside:
                    occurrence = inside[0]
                    context_verified = True
                else:
                    context_note = "给出的上下文里并没有该错别字，位置按错别字本身推算"

        verdict = _tender_verdict(wrong, normalized_tender, skeleton_tender)
        if verdict != "none":
            status = "needs_review"
            reason = "该文本出现在招标原文里，可能是招标方既有用词，需人工确认后再改"
            if verdict == "skeleton":
                reason = "该文本与招标原文用词一致（仅数字不同），可能不是错别字，需人工确认"
        else:
            status = "confirmed"
            reason = "已在原文定位，且未出现在招标原文中"

        findings.append({
            "id": item["id"], "status": status, "status_label": TYPO_STATUS_LABELS[status],
            "wrong": wrong, "right": right, "note": item["note"],
            "location": f"{occurrence['line']}:{occurrence['column']}",
            "line": occurrence["line"], "column": occurrence["column"],
            "line_text": occurrence["line_text"],
            "occurrences": len(hits),
            "context_verified": context_verified,
            "tender_verdict": verdict,
            "reason": reason + (f"；{context_note}" if context_note else ""),
        })

    order = {"confirmed": 0, "needs_review": 1}
    # 按**文档顺序**输出（逐行看一遍最自然）；同一位置下「确认」排在「待核对」前面。
    # 需要按状态分组时，xlsx 里有「状态」列可直接筛选/排序。
    findings.sort(key=lambda row: (row["line"], row["column"], order.get(row["status"], 9)))
    counts = {status: sum(1 for row in findings + unlocated + invalid if row.get("status") == status)
              for status in TYPO_STATUS}
    return {
        "kind": "typo", "source_name": resolved_name, "source_chars": len(source),
        "findings": findings, "unlocated": unlocated, "invalid": invalid,
        "summary": {"candidates": len(items), **counts,
                    "tender_used": bool(normalized_tender),
                    "ok": not findings},
        "warnings": ([] if normalized_tender else
                     ["未提供招标原文：无法判断「错别字」是否是招标方既有用词，"
                      "建议传 tender_text 以降低误报。"]),
        "note": ("错别字候选均已**回原文定位**：定位不到的判为疑似误判并排除（防模型造错）；"
                 "出现在招标原文里的降级为「待核对」。建议以「确认」项为准，其余需人工过一遍。"),
    }


# ---------- 逻辑谬误 ----------


def check_logic(segments: Any, extractions: Any, facts_id: str | None = None) -> dict:
    """逐段上报事实 → 归集登记 → 跨段取值不一致即判前后矛盾（复用一致性台账）。"""
    if not isinstance(segments, list) or not segments:
        raise ReviewError("segments 必须是非空数组（每个待检查的小节/片段）")
    ledger = bid_consistency.build_ledger(
        {"sections": segments, "extractions": extractions, "facts_id": facts_id})

    findings: list[dict] = []
    for conflict in ledger.get("conflicts", []):
        titles = []
        for value in conflict.get("values", []):
            title = value.get("section_title") or value.get("section_id") or ""
            if title and title not in titles:
                titles.append(title)
        findings.append({
            "id": f"c-{conflict.get('name', '')[:20]}",
            "type": "前后矛盾", "category": conflict.get("category") or "其他",
            "name": conflict.get("name") or "",
            "segments": "、".join(titles),
            "distinct": " / ".join(conflict.get("distinct") or []),
            "detail": conflict.get("message") or "",
        })
    for issue in ledger.get("issues", []):
        findings.append({
            "id": issue.get("section_id") or "", "type": "逻辑谬误",
            "category": bid_consistency.ISSUE_CLASS_LABELS.get(issue.get("kind"), "其他"),
            "name": issue.get("detail") or "",
            "segments": issue.get("section_title") or issue.get("section_id") or "",
            "distinct": "", "detail": issue.get("detail") or "",
        })

    register = [{"category": row.get("category"), "name": row.get("name"),
                 "values": [value.get("value") for value in row.get("values", [])],
                 "conflict": bool(row.get("conflict"))}
                for row in ledger.get("fact_table", [])]
    return {"kind": "logic", "ledger_id": ledger.get("ledger_id"), "findings": findings,
            "register": register, "summary": ledger.get("summary"),
            "warnings": ([] if register else ["没有可归集的事实：请把各段的关键事实（人数/工期/金额/"
                                              "技术参数等）一并上报，否则无法做前后矛盾检测。"]),
            "note": ("事实登记表按 (分类, 名称) 归集后比对：**跨段取值不一致即判前后矛盾**；"
                     "模型自报的「小节内部矛盾」也一并列出。用词/称谓/文风差异不参与比较。")}


# ---------- 与废标项检查共用导出 ----------


def sheets_for(typo_report: dict | None = None, logic_report: dict | None = None) -> tuple[list, list]:
    """把两类报告整理成 `build_workbook` 的 (extra_sheets, extra_overview_rows)。"""
    sheets: list[dict] = []
    overview: list[tuple] = []

    if typo_report:
        rows = [[index, row.get("location") or "", row.get("line_text") or "", row.get("wrong") or "",
                 row.get("right") or "", row.get("reason") or "", row.get("status_label") or "",
                 row.get("occurrences") or 0, row.get("id") or ""]
                for index, row in enumerate(typo_report.get("findings") or [], 1)]
        sheets.append({"title": "错别字", "columns": TYPO_COLUMNS, "rows": rows})
        summary = typo_report.get("summary") or {}
        overview.extend([("错别字-候选数", summary.get("candidates", 0)),
                         ("错别字-已确认", summary.get("confirmed", 0)),
                         ("错别字-待核对", summary.get("needs_review", 0)),
                         ("错别字-未定位（已排除）", summary.get("unlocated", 0))])

    if logic_report:
        rows = [[index, row.get("type") or "", row.get("category") or "", row.get("name") or "",
                 row.get("segments") or "", row.get("distinct") or "", row.get("detail") or "",
                 row.get("id") or ""]
                for index, row in enumerate(logic_report.get("findings") or [], 1)]
        sheets.append({"title": "逻辑问题", "columns": LOGIC_COLUMNS, "rows": rows})
        summary = logic_report.get("summary") or {}
        overview.extend([("逻辑-checked小节", summary.get("sections", 0)),
                         ("逻辑-前后矛盾", summary.get("conflicts", 0)),
                         ("逻辑-内部矛盾", summary.get("internal", 0))])
    return sheets, overview


# ---------- 结果落盘（供导出时取用，避免来回传大 JSON） ----------


def review_dir() -> Path:
    return Path(settings.bid_state_dir) / "review"


def _review_path(review_id: str) -> Path:
    if not REVIEW_ID.fullmatch(str(review_id or "")):
        raise ReviewError(f"review_id 非法：{review_id!r}")
    return review_dir() / f"{review_id}.json"


def persist_review(report: dict) -> dict:
    review_id = f"rev-{time.strftime('%Y%m%d')}-{uuid.uuid4().hex[:8]}"
    record = {"version": STATE_VERSION, "review_id": review_id, "created_at": _now(), **report}
    path = _review_path(review_id)
    path.parent.mkdir(parents=True, exist_ok=True)
    handle_fd, temporary = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    try:
        with os.fdopen(handle_fd, "w", encoding="utf-8") as handle:
            handle.write(json.dumps(record, ensure_ascii=False, indent=2))
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)
    record["path"] = str(path)
    return record


def load_review(review_id: str, expect_kind: str | None = None) -> dict:
    path = _review_path(review_id)
    if not path.is_file():
        raise ReviewError(f"未找到检查结果：{review_id}")
    try:
        record = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise ReviewError(f"检查结果无法解析：{type(exc).__name__}") from exc
    if not isinstance(record, dict) or int(record.get("version") or 0) != STATE_VERSION:
        raise ReviewError("检查结果版本不匹配，请重新检查")
    if expect_kind and record.get("kind") != expect_kind:
        raise ReviewError(f"{review_id} 不是 {expect_kind} 结果（实际：{record.get('kind')}）")
    return record


def run_typo_check(source_text: Any = None, candidates: Any = None, tender_text: Any = None,
                   source_name: str = "", source_path: Any = None, tender_path: Any = None) -> dict:
    """回验错别字并落盘，返回带 review_id 的报告。"""
    return persist_review(check_typos(source_text, candidates, tender_text, source_name,
                                     source_path, tender_path))


def run_logic_check(segments: Any, extractions: Any, facts_id: str | None = None) -> dict:
    """逻辑检查并落盘，返回带 review_id 的报告。"""
    return persist_review(check_logic(segments, extractions, facts_id))


def export_combined(*, rejection_session_id: str | None = None, typo_id: str | None = None,
                    logic_id: str | None = None, name: str | None = None,
                    project_name: str = "") -> dict:
    """把废标项 / 错别字 / 逻辑谬误的检查结果导出到**同一份 xlsx**。"""
    report: dict = {}
    if rejection_session_id:
        state = bid_rejection.load_state(rejection_session_id)
        report = state.get("report") or bid_rejection.finalize(state)

    typo_report = load_review(typo_id, "typo") if typo_id else None
    logic_report = load_review(logic_id, "logic") if logic_id else None
    sheets, overview = sheets_for(typo_report, logic_report)
    if not report and not sheets:
        raise ReviewError("没有可导出的检查结果：请至少提供 rejection_session_id / typo_id / logic_id 之一")

    if not report:
        report = {"session_id": "", "project": {"name": _clip(project_name, 120)},
                  "summary": {}, "warnings": ["本次导出不含废标项检查结果（仅错别字 / 逻辑谬误）。"]}
    if project_name and not (report.get("project") or {}).get("name"):
        report.setdefault("project", {})["name"] = _clip(project_name, 120)

    default_name = (f"{rejection_session_id}-检查报告.xlsx" if rejection_session_id
                    else f"正文检查报告-{time.strftime('%Y%m%d-%H%M%S')}.xlsx")
    result = bid_rejection.export_workbook(report, extra_sheets=sheets, extra_overview_rows=overview,
                                          name=name or default_name)
    result["sheets"] = [sheet["title"] for sheet in sheets]
    return result


__all__ = ["ReviewError", "TYPO_STATUS", "TYPO_STATUS_LABELS", "TYPO_COLUMNS", "LOGIC_COLUMNS",
           "locate_text", "check_typos", "check_logic", "sheets_for", "persist_review",
           "load_review", "run_typo_check", "run_logic_check", "export_combined", "review_dir"]
