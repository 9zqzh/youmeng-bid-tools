"""一致性审计：程序做台账与冲突判定，模型只逐节上报事实与问题。

为什么这样拆：把「比较」交给程序，才能可靠、可复现，也才可能覆盖几十节的长文档。
模型只并发核对单个小节，上报本节的**类型化事实取值**（按固定分类）与两类问题
（与全局事实冲突 / 小节内部矛盾）；程序把各节事实按 (分类, 名称) 归集成台账，找出
**跨节同名不同取值**，以及与**全局事实清单**不一致的取值。

**去噪是设计出来的，不是靠提示词**：只比较类型化事实，用词、称谓、文风、篇幅、
承诺语气强弱根本不会进入比较，因此不会产生噪音问题。

台账以 JSON 落盘（`BID_OUTPUT_DIR/ledger/<ledger_id>.json`），走「先建账、再定口径、
后派发修复」三步。
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
from . import bid_facts


class ConsistencyError(ValueError):
    """一致性审计模块的可预期错误（对应 HTTP 422）。"""


LEDGER_VERSION = 1
ISSUE_CLASSES = ("against_global", "internal", "other")
ISSUE_CLASS_LABELS = {"against_global": "与全局事实冲突", "internal": "小节内部矛盾", "other": "其他"}

_LEDGER_ID = re.compile(r"^led-\d{8}-[a-f0-9]{8}$")
_CONFLICT_ORDER = {"against_global": 0, "cross_section": 1}


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _clip(value: Any, limit: int) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()[:limit]


def _classify_issue(raw_type: str) -> str:
    text = str(raw_type or "")
    lowered = text.lower()
    if "全局" in text or "global" in lowered:
        return "against_global"
    if "内部" in text or "矛盾" in text or "internal" in lowered:
        return "internal"
    return "other"


# ---------- 输入归一化 ----------


def normalize_sections(raw: Any) -> list[dict]:
    """归一化小节清单；缺 id 时按顺序分配稳定 id。"""
    if raw is None:
        return []
    if not isinstance(raw, list):
        raise ConsistencyError("sections 必须是数组")
    sections: list[dict] = []
    seen: set[str] = set()
    for index, item in enumerate(raw, 1):
        if isinstance(item, str):
            item = {"title": item}
        if not isinstance(item, dict):
            continue
        section_id = _clip(item.get("id") or item.get("section_id"), 40) or f"sec_{index:04d}"
        if section_id in seen:
            raise ConsistencyError(f"小节 id 重复：{section_id}")
        seen.add(section_id)
        sections.append({"id": section_id, "title": _clip(item.get("title"), 120)})
    return sections


def normalize_extractions(raw: Any, section_ids: set[str]) -> list[dict]:
    """归一化模型逐节上报的抽取结果。"""
    if raw is None:
        return []
    if not isinstance(raw, list):
        raise ConsistencyError("extractions 必须是数组")
    normalized: list[dict] = []
    unknown: set[str] = set()
    for entry in raw:
        if not isinstance(entry, dict):
            continue
        section_id = _clip(entry.get("section_id") or entry.get("id"), 40)
        if section_ids and section_id not in section_ids:
            unknown.add(section_id)
            continue
        facts: list[dict] = []
        for fact in entry.get("facts") or []:
            if not isinstance(fact, dict):
                continue
            name = _clip(fact.get("name") or fact.get("key"), 80)
            if not name:
                continue
            category = _clip(fact.get("category"), 20)
            if category not in bid_facts.FACT_CATEGORIES:
                category = "其他"
            facts.append({"category": category, "name": name,
                          "value": _clip(fact.get("value"), 400),
                          "quote": _clip(fact.get("quote"), 400)})
        issues: list[dict] = []
        for issue in entry.get("issues") or []:
            if isinstance(issue, str):
                issue = {"detail": issue}
            if not isinstance(issue, dict):
                continue
            kind = _classify_issue(issue.get("type") or issue.get("kind"))
            detail = _clip(issue.get("detail") or issue.get("message"), 400)
            if not detail:
                continue
            issues.append({"kind": kind, "detail": detail,
                           "quote": _clip(issue.get("quote"), 400)})
        normalized.append({"section_id": section_id, "facts": facts, "issues": issues})
    if unknown:
        raise ConsistencyError("extractions 引用了未知小节 id：" + "、".join(sorted(unknown))[:200])
    return normalized


# ---------- 台账 ----------


def build_ledger(payload: dict) -> dict:
    """归集各节事实 → 找出跨节冲突与违反全局事实的取值。"""
    if not isinstance(payload, dict):
        raise ConsistencyError("payload 必须是对象")
    sections = normalize_sections(payload.get("sections"))
    section_index = {section["id"]: section for section in sections}
    extractions = normalize_extractions(payload.get("extractions"), set(section_index))

    global_by_key: dict[str, dict] = {}
    facts_id = _clip(payload.get("facts_id"), 40)
    if facts_id:
        record = bid_facts.load_facts(facts_id)
        for fact in record.get("facts", []):
            global_by_key[bid_facts.fact_key(fact["name"])] = fact

    table: dict[str, dict] = {}
    for entry in extractions:
        for fact in entry["facts"]:
            key = bid_facts.fact_key(fact["name"])
            row = table.setdefault(key, {"name": fact["name"], "category": fact["category"], "values": []})
            if row["category"] == "其他" and fact["category"] != "其他":
                row["category"] = fact["category"]
            row["values"].append({"section_id": entry["section_id"],
                                  "section_title": section_index.get(entry["section_id"], {}).get("title", ""),
                                  "value": fact["value"], "quote": fact["quote"]})

    conflicts: list[dict] = []
    for key, row in table.items():
        distinct: list[tuple[str, str]] = []
        for item in row["values"]:
            value_key = bid_facts.value_key(item["value"])
            if value_key and value_key not in [seen for seen, _ in distinct]:
                distinct.append((value_key, item["value"]))
        row["distinct"] = [value for _, value in distinct]
        row["conflict"] = len(distinct) > 1

        global_fact = global_by_key.get(key)
        mismatched: list[dict] = []
        if global_fact and global_fact.get("value"):
            expected = bid_facts.value_key(global_fact["value"])
            mismatched = [item for item in row["values"]
                          if item["value"] and bid_facts.value_key(item["value"]) != expected]
        row["against_global"] = global_fact["value"] if mismatched else None
        row["against_global_locked"] = bool(global_fact.get("locked")) if mismatched else False

        # 同一个事实只报一条冲突：违反全局事实优先（更权威），并标注是否同时跨节不一致。
        if mismatched:
            kind = "against_global"
            message = "《{}》与全局事实（{}）不一致：{}".format(
                row["name"], global_fact["value"], " / ".join(item["value"] for item in mismatched))
        elif row["conflict"]:
            kind = "cross_section"
            message = "《{}》在不同小节取值不一致：{}".format(row["name"], " / ".join(row["distinct"]))
        else:
            continue

        conflicts.append({
            "kind": kind, "category": row["category"], "name": row["name"],
            "values": row["values"], "distinct": row["distinct"],
            "global_value": global_fact["value"] if mismatched else None,
            "locked": bool(global_fact.get("locked")) if mismatched else False,
            "also_cross_section": bool(row["conflict"] and mismatched),
            "message": message,
        })

    def _conflict_sort(item: dict):
        return (_CONFLICT_ORDER.get(item["kind"], 9), 0 if item.get("locked") else 1, item["name"])

    conflicts.sort(key=_conflict_sort)

    issues: list[dict] = []
    for entry in extractions:
        for issue in entry["issues"]:
            issues.append({**issue, "section_id": entry["section_id"],
                           "section_title": section_index.get(entry["section_id"], {}).get("title", "")})
    issues.sort(key=lambda item: (item["kind"], item["section_id"]))

    fact_table = [table[key] for key in sorted(table, key=lambda k: (
        bid_facts.FACT_CATEGORIES.index(table[k]["category"]) if table[k]["category"] in bid_facts.FACT_CATEGORIES else 99,
        table[k]["name"]))]

    ledger = {
        "version": LEDGER_VERSION,
        "ledger_id": f"led-{time.strftime('%Y%m%d')}-{uuid.uuid4().hex[:8]}",
        "created_at": _now(),
        "facts_id": facts_id or None,
        "sections": sections,
        "fact_table": fact_table,
        "conflicts": conflicts,
        "issues": issues,
        "summary": {
            "sections": len(sections),
            "facts": len(fact_table),
            "conflicts": len(conflicts),
            "cross_section": sum(1 for item in conflicts if item["kind"] == "cross_section"),
            "against_global": sum(1 for item in conflicts if item["kind"] == "against_global"),
            "also_cross_section": sum(1 for item in conflicts if item.get("also_cross_section")),
            "locked_violations": sum(1 for item in conflicts if item.get("locked")),
            "internal": sum(1 for item in issues if item["kind"] == "internal"),
            "against_global_reported": sum(1 for item in issues if item["kind"] == "against_global"),
            "other_issues": sum(1 for item in issues if item["kind"] == "other"),
            "ok": not conflicts and not issues,
        },
    }
    save_ledger(ledger)
    return ledger


# ---------- 修复派发 ----------


def normalize_rules(raw: Any) -> list[dict]:
    if raw is None:
        return []
    if not isinstance(raw, list):
        raise ConsistencyError("rules 必须是数组")
    rules: list[dict] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        target = _clip(item.get("target") or item.get("name"), 80)
        value = _clip(item.get("value"), 400)
        if not target or not value:
            continue
        rules.append({"target": target, "value": value, "note": _clip(item.get("note"), 200)})
    return rules


def repair_plan(ledger_id: str, rules: Any) -> dict:
    """按统一口径生成逐节修复任务，并报告未被口径覆盖的冲突。"""
    ledger = load_ledger(ledger_id)
    section_index = {section["id"]: section for section in ledger.get("sections", [])}
    rules = normalize_rules(rules)
    by_target = {bid_facts.fact_key(rule["target"]): rule for rule in rules}

    tasks: dict[str, dict] = {}

    def add_task(section_id: str, instruction: str) -> None:
        task = tasks.setdefault(section_id, {
            "section_id": section_id,
            "title": section_index.get(section_id, {}).get("title", ""),
            "instructions": []})
        task["instructions"].append(instruction)

    applied: list[dict] = []
    unresolved: list[dict] = []
    for conflict in ledger.get("conflicts", []):
        rule = by_target.get(bid_facts.fact_key(conflict["name"]))
        if rule is None:
            unresolved.append(conflict)
            continue
        target_key = bid_facts.value_key(rule["value"])
        # 只对「取值与统一口径不同」的小节下发指令，且同一小节同一事实只发一次。
        changed_by_section: dict[str, str] = {}
        for item in conflict["values"]:
            if bid_facts.value_key(item["value"]) == target_key:
                continue
            changed_by_section.setdefault(item["section_id"], item["value"])
        applied.append({"target": conflict["name"], "value": rule["value"], "kind": conflict["kind"],
                        "sections": sorted(changed_by_section),
                        "already_consistent": not changed_by_section})
        for section_id, current in sorted(changed_by_section.items()):
            add_task(section_id,
                     "把《{}》统一为「{}」（本节当前值：{}）".format(conflict["name"], rule["value"], current))

    for issue in ledger.get("issues", []):
        if issue["kind"] == "internal":
            add_task(issue["section_id"], "修复本节内部矛盾：" + issue["detail"])
        else:
            unresolved.append({"kind": "issue", "name": issue.get("detail"), "category": "其他",
                               "values": [{"section_id": issue["section_id"], "value": issue["detail"]}],
                               "distinct": [issue["detail"]], "message": issue["detail"]})

    used = {bid_facts.fact_key(item["target"]) for item in applied}
    unused_rules = [rule for key, rule in by_target.items() if key not in used]

    task_list = [tasks[key] for key in sorted(tasks)]
    summary = {
        "conflicts": len(ledger.get("conflicts", [])),
        "resolved_conflicts": len(applied),
        "unresolved_conflicts": len(unresolved),
        "sections_to_fix": len(task_list),
        "instructions": sum(len(task["instructions"]) for task in task_list),
        "ok": not unresolved,
    }
    # 把修复计划摘要写回台账：一是审计留痕，二是供项目工作流校验关卡。
    ledger["repair"] = {"at": _now(), "rules": rules, "summary": summary,
                        "unresolved": [{"kind": item.get("kind"), "name": item.get("name")} for item in unresolved]}
    save_ledger(ledger)
    return {
        "ledger_id": ledger_id,
        "rules": rules,
        "applied_rules": applied,
        "tasks": task_list,
        "unresolved": unresolved,
        "unused_rules": unused_rules,
        "summary": summary,
    }


# ---------- 存取 ----------


def ledger_dir() -> Path:
    return Path(settings.bid_state_dir) / "ledger"


def _ledger_path(ledger_id: str) -> Path:
    if not _LEDGER_ID.fullmatch(str(ledger_id or "")):
        raise ConsistencyError(f"ledger_id 非法：{ledger_id!r}")
    return ledger_dir() / f"{ledger_id}.json"


def save_ledger(ledger: dict) -> Path:
    path = _ledger_path(str(ledger["ledger_id"]))
    path.parent.mkdir(parents=True, exist_ok=True)
    handle_fd, temporary = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    try:
        with os.fdopen(handle_fd, "w", encoding="utf-8") as handle:
            handle.write(json.dumps(ledger, ensure_ascii=False, indent=2))
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)
    return path


def load_ledger(ledger_id: str) -> dict:
    path = _ledger_path(ledger_id)
    if not path.is_file():
        raise ConsistencyError(f"未找到一致性台账：{ledger_id}")
    try:
        ledger = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise ConsistencyError(f"台账无法解析：{type(exc).__name__}") from exc
    if not isinstance(ledger, dict) or int(ledger.get("version") or 0) != LEDGER_VERSION:
        raise ConsistencyError("台账版本不匹配，请重新建账")
    return ledger


__all__ = ["ConsistencyError", "ISSUE_CLASSES", "ISSUE_CLASS_LABELS", "normalize_sections",
           "normalize_extractions", "normalize_rules", "build_ledger", "repair_plan",
           "save_ledger", "load_ledger", "ledger_dir"]
