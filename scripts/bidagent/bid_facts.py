"""全局事实设定：在正文生成前固化「跨章节反复出现的事实」，避免各节随机漂移。

问题：技术方案 / 商务标动辄几十节，若每节各自生成，人员数量、工期、质保期、
型号、金额等会互相矛盾。做法是先产出一份**权威事实清单**，逐节完整注入。

本模块只做**确定性**的部分（不调用模型）：
- `seed_company_facts`：从企业资质清单注入主体信息（名称 / 统一社会信用代码 /
  法定代表人），这部分是**锁定事实**，模型不得改写。
- `normalize_facts`：归一化模型提交的事实（分类、去空、同名去重）。
- `lint_facts`：检查占位符残留、同名冲突、与锁定事实冲突。
- `render_facts`：渲染成**逐字节一致**的注入块，保证每节看到同一份事实。

清单以 JSON 落盘（`BID_OUTPUT_DIR/facts/<facts_id>.json`），供 draft 阶段复用。
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


class FactsError(ValueError):
    """全局事实模块的可预期错误（对应 HTTP 422）。"""


STATE_VERSION = 1

# 与一致性审计共用的事实分类，便于台账按同一维度归集。
FACT_CATEGORIES = ("企业主体", "人员与数量", "日期与期限", "频次与时限", "地点与范围",
                   "金额", "技术参数", "编号与名称", "责任归属", "其他")

# 占位符 / 未落定标记：出现即视为「事实未确认」，不能靠它写正文。
_PLACEHOLDER_PATTERN = re.compile(
    r"【待[填写补充确认核实]+】|（待[填写补充确认核实]+）|\(待[填写补充确认核实]+\)|"
    r"待填写|待补充|待确认|待核实|\bTBD\b|\bTODO\b|X{2,}|x{3,}"
)

_FACTS_ID = re.compile(r"^facts-\d{8}-[a-f0-9]{8}$")


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _clip(value: Any, limit: int) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()[:limit]


def fact_key(name: str) -> str:
    """事实名归一化 key（用于判同名 / 冲突）。"""
    return re.sub(r"[\s　:：=＝\-—_（）()【】\[\]]+", "", str(name or "")).lower()


def value_key(value: str) -> str:
    """事实值归一化 key（用于判冲突）。全角转半角、去空白与常见标点。"""
    text = str(value or "")
    converted = []
    for char in text:
        code = ord(char)
        if code == 0x3000:
            converted.append(" ")
        elif 0xFF01 <= code <= 0xFF5E:
            converted.append(chr(code - 0xFEE0))
        else:
            converted.append(char)
    text = "".join(converted)
    return re.sub(r"[\s,，、;；.。:：/\\|]+", "", text).lower()


def has_placeholder(text: str) -> bool:
    return bool(_PLACEHOLDER_PATTERN.search(str(text or "")))


# ---------- 归一化 ----------


def normalize_facts(raw: Any, *, source: str = "model") -> list[dict]:
    """归一化事实清单：分类归位、去空、同名去重（后者保留第一个）。"""
    if raw is None:
        return []
    if isinstance(raw, dict):
        raw = raw.get("facts") if isinstance(raw.get("facts"), list) else [raw]
    if not isinstance(raw, list):
        raise FactsError("facts 必须是数组或含 facts 数组的对象")

    facts: list[dict] = []
    seen: set[str] = set()
    for item in raw:
        if isinstance(item, str):
            item = {"name": item}
        if not isinstance(item, dict):
            continue
        name = _clip(item.get("name") or item.get("key") or item.get("title"), 80)
        value = _clip(item.get("value") or item.get("content"), 400)
        if not name:
            continue
        key = fact_key(name)
        if key in seen:
            continue
        seen.add(key)
        category = _clip(item.get("category"), 20)
        if category not in FACT_CATEGORIES:
            category = "其他"
        facts.append({
            "category": category,
            "name": name,
            "value": value,
            "source": _clip(item.get("source") or source, 60),
            "locked": bool(item.get("locked")),
        })
    return facts


def seed_company_facts(company: str, profile: dict | None = None) -> list[dict]:
    """从资质清单注入企业主体信息，作为**锁定事实**，模型不得改写。"""
    profile = profile if isinstance(profile, dict) else {}
    seeded: list[dict] = []
    pairs = [
        ("投标主体全称", _clip(company, 120), "资质清单 meta.company"),
        ("统一社会信用代码", _clip(profile.get("unified_social_credit_code"), 40),
         "资质清单 meta.company_profile"),
        ("法定代表人", _clip(profile.get("legal_representative"), 40), "资质清单 meta.company_profile"),
    ]
    for name, value, source in pairs:
        if not value:
            continue
        seeded.append({"category": "企业主体", "name": name, "value": value,
                       "source": source, "locked": True})
    return seeded


# ---------- 校验 ----------


def lint_facts(facts: list[dict]) -> list[dict]:
    """检查占位符残留与同名冲突，返回问题列表（不抛错）。"""
    issues: list[dict] = []
    by_key: dict[str, dict] = {}
    for fact in facts:
        key = fact_key(fact.get("name", ""))
        value = str(fact.get("value") or "")
        if not value:
            issues.append({"level": "warning", "kind": "empty_value",
                           "name": fact.get("name"), "message": f"「{fact.get('name')}」的值为空，相当于没有设定"})
        elif has_placeholder(value):
            issues.append({"level": "warning", "kind": "placeholder",
                           "name": fact.get("name"), "value": value,
                           "message": f"「{fact.get('name')}」仍含占位符，须先改为真实数据再生成正文"})
        previous = by_key.get(key)
        if previous is None:
            by_key[key] = fact
            continue
        if value_key(previous.get("value", "")) != value_key(value):
            # 锁定事实（来自资质清单）优先
            winner, loser = (previous, fact) if previous.get("locked") else (fact, previous)
            issues.append({"level": "error", "kind": "duplicate_conflict",
                           "name": fact.get("name"),
                           "values": [previous.get("value"), value],
                           "keep": winner.get("value"),
                           "message": f"「{fact.get('name')}」出现两个不同取值：{previous.get('value')} / {value}"})
    return issues


def merge_with_locked(base: list[dict], incoming: list[dict]) -> tuple[list[dict], list[dict]]:
    """合并事实：锁定事实（资质清单）优先，返回 (合并结果, 冲突问题)。"""
    merged: list[dict] = list(base)
    issues: list[dict] = []
    index = {fact_key(fact["name"]): fact for fact in merged}
    for fact in incoming:
        key = fact_key(fact["name"])
        current = index.get(key)
        if current is None:
            merged.append(fact)
            index[key] = fact
            continue
        if value_key(current.get("value", "")) == value_key(fact.get("value", "")):
            continue
        if current.get("locked"):
            issues.append({"level": "error", "kind": "locked_conflict", "name": fact["name"],
                           "values": [current.get("value"), fact.get("value")],
                           "keep": current.get("value"),
                           "message": (f"「{fact['name']}」与资质清单主体信息冲突："
                                       f"清单为「{current.get('value')}」，提交为「{fact.get('value')}」，已以清单为准")})
        else:
            current["value"] = fact.get("value")
            issues.append({"level": "warning", "kind": "overridden", "name": fact["name"],
                           "values": [current.get("value"), fact.get("value")],
                           "message": f"「{fact['name']}」重复提交，已采用后者"})
    return merged, issues


# ---------- 渲染 ----------


def render_facts(facts: list[dict]) -> str:
    """把事实渲染成逐字节一致的注入块（按固定分类顺序）。"""
    if not facts:
        return ""
    lines = ["# 全局事实（各节必须与本表一致；表中标注的待办项不得凭空编造）", ""]
    for category in FACT_CATEGORIES:
        group = [fact for fact in facts if fact.get("category") == category]
        if not group:
            continue
        lines.append(f"## {category}")
        for fact in group:
            flag = "（锁定）" if fact.get("locked") else ""
            value = fact.get("value") or "（未设定）"
            lines.append(f"- {fact['name']}{flag}：{value}")
        lines.append("")
    return "\n".join(lines).strip() + "\n"


# ---------- 存取 ----------


def facts_dir() -> Path:
    return Path(settings.bid_state_dir) / "facts"


def _facts_path(facts_id: str) -> Path:
    if not _FACTS_ID.fullmatch(str(facts_id or "")):
        raise FactsError(f"facts_id 非法：{facts_id!r}")
    return facts_dir() / f"{facts_id}.json"


def save_facts(record: dict) -> Path:
    path = _facts_path(str(record["facts_id"]))
    path.parent.mkdir(parents=True, exist_ok=True)
    record["updated_at"] = _now()
    handle_fd, temporary = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    try:
        with os.fdopen(handle_fd, "w", encoding="utf-8") as handle:
            handle.write(json.dumps(record, ensure_ascii=False, indent=2))
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)
    return path


def load_facts(facts_id: str) -> dict:
    path = _facts_path(facts_id)
    if not path.is_file():
        raise FactsError(f"未找到事实清单：{facts_id}")
    try:
        record = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise FactsError(f"事实清单无法解析：{type(exc).__name__}") from exc
    if not isinstance(record, dict) or int(record.get("version") or 0) != STATE_VERSION:
        raise FactsError("事实清单版本不匹配，请重新初始化")
    return record


# ---------- 对外流程 ----------


def init_facts(project: dict | None, facts: Any, manifest: dict | None = None) -> dict:
    """建档：注入资质清单主体信息 → 合并模型事实 → 校验 → 落盘。"""
    project = project if isinstance(project, dict) else {}
    seeded: list[dict] = []
    if manifest:
        seeded = seed_company_facts(manifest.get("company") or "", manifest.get("profile"))
    merged, conflicts = merge_with_locked(seeded, normalize_facts(facts))
    facts_id = f"facts-{time.strftime('%Y%m%d')}-{uuid.uuid4().hex[:8]}"
    record = {
        "version": STATE_VERSION,
        "facts_id": facts_id,
        "project": {"name": _clip(project.get("name") or project.get("project_name"), 120),
                    "bid_no": _clip(project.get("bid_no"), 60)},
        "created_at": _now(),
        "updated_at": _now(),
        "facts": merged,
    }
    issues = conflicts + lint_facts(merged)
    record["issues"] = issues
    save_facts(record)
    return {"facts_id": facts_id, "project": record["project"], "facts": merged,
            "issues": issues, "rendered": render_facts(merged),
            "ok": not any(item["level"] == "error" for item in issues)}


def update_facts(facts_id: str, facts: Any, manifest: dict | None = None) -> dict:
    """按 id 全量替换事实（人工评审后回填），重新校验并渲染。"""
    record = load_facts(facts_id)
    seeded = seed_company_facts(manifest.get("company") or "", manifest.get("profile")) if manifest else []
    locked = [fact for fact in record.get("facts", []) if fact.get("locked")] if not seeded else seeded
    merged, conflicts = merge_with_locked(locked, normalize_facts(facts))
    record["facts"] = merged
    issues = conflicts + lint_facts(merged)
    record["issues"] = issues
    save_facts(record)
    return {"facts_id": facts_id, "project": record.get("project", {}), "facts": merged,
            "issues": issues, "rendered": render_facts(merged),
            "ok": not any(item["level"] == "error" for item in issues)}


def get_rendered(facts_id: str) -> dict:
    record = load_facts(facts_id)
    return {"facts_id": facts_id, "rendered": render_facts(record.get("facts", [])),
            "count": len(record.get("facts", []))}


__all__ = ["FactsError", "FACT_CATEGORIES", "has_placeholder", "normalize_facts", "seed_company_facts",
           "lint_facts", "merge_with_locked", "render_facts", "init_facts", "update_facts",
           "get_rendered", "load_facts", "save_facts", "facts_dir", "fact_key", "value_key"]
