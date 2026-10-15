"""项目工作流：把散落的 facts_id / ledger_id / session_id 收进一个「项目」，并做关卡校验。

要解决的问题：整条链路是「招标解读 → 全局事实 → 正文 → 一致性审计 → 两册 Word → 废标项检查」，
每个阶段产物是一个 id，此前全靠 Agent 自己记、自己保证顺序。一旦漏做「全局事实」就直接写正文，
各节数据必然互相打架；而这类错误在最终成品里才暴露，代价很高。

本模块把这条链路变成一个**带关卡的项目**：
- 阶段产物（artifacts）集中挂在项目上，Agent 只需带一个 `project_id`；
- `advance` 前会**真的去校验**产物（不是只记个标记）：
  事实清单无 error 级问题、台账已给出统一口径且无未解决冲突、Word 文件确实存在、
  废标项会话确实跑完；
- 关卡不过就拒绝推进，并给出**具体是哪一条不满足**；确有必要可 `force` 跳过并留痕。

本模块不调用模型，只做编排与校验。状态落盘在 `BID_OUTPUT_DIR/project/<project_id>.json`。
"""
from __future__ import annotations

import hashlib
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
from . import bid_consistency, bid_facts, bid_rejection


class ProjectError(ValueError):
    """项目工作流的可预期错误（对应 HTTP 422）。"""


STATE_VERSION = 1

# 阶段顺序即依赖顺序；artifacts 是「进入该阶段必须带来的产物」。
STAGES: list[dict] = [
    {"key": "analysis", "label": "招标解读与资质比对", "artifacts": [],
     "guide": "用 list / match 完成招标解读与资质比对，汇总门槛缺口与待补充项。"},
    {"key": "facts", "label": "全局事实设定", "artifacts": ["facts_id"],
     "guide": "用 facts-init 建档（--facts @facts.json），处理掉 issues 里的占位符与冲突。"},
    {"key": "content", "label": "正文撰写", "artifacts": [],
     "guide": "用 facts-render 取事实块并**完整注入**，逐节撰写正文。"},
    {"key": "consistency", "label": "一致性审计", "artifacts": ["ledger_id"],
     "guide": "逐节抽取事实与问题 → cons-build 建账；有冲突则 cons-repair 给出统一口径，按 tasks 修正各节。"},
    {"key": "draft", "label": "两册 Word 生成", "artifacts": ["files"],
     "guide": "用 draft 生成两册 Word（--sections @sections.json [--qualifications ...]），把链接交给用户。"},
    {"key": "rejection", "label": "废标项检查", "artifacts": ["rejection_session_id"],
     "guide": "用 reject-init 建档 → 逐片段 reject-apply → reject-finalize → reject-export。"},
]

STAGE_KEYS = [stage["key"] for stage in STAGES]
_STAGE_INDEX = {stage["key"]: index for index, stage in enumerate(STAGES)}
_PROJECT_ID = re.compile(r"^prj-\d{8}-[a-f0-9]{8}$")
_ARTIFACT_NAMES = {name for stage in STAGES for name in stage["artifacts"]}


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _clip(value: Any, limit: int) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()[:limit]


# ---------- 产物校验（关卡的核心） ----------


def _check_artifact(name: str, value: Any) -> tuple[list[str], list[str]]:
    """严格校验单个产物。返回 (blockers, warnings)。

    规则：产物必须**真的达标**，不只是「有个 id」——
    事实清单无 error 级问题、台账已给出统一口径且无未解决冲突、
    Word 文件确实在盘、废标项会话确实跑完。
    """
    blockers: list[str] = []
    warnings: list[str] = []
    if not value:
        blockers.append(f"缺少产物 {name}：请先完成对应阶段")
        return blockers, warnings

    if name == "facts_id":
        try:
            record = bid_facts.load_facts(str(value))
        except ValueError as exc:
            blockers.append(f"facts_id 无效：{exc}")
            return blockers, warnings
        facts_list = record.get("facts", [])
        if not facts_list:
            blockers.append("事实清单为空：请先固化若干跨节事实（人员数 / 工期 / 质保期 / 金额 / 技术参数等）")
            return blockers, warnings
        issues = bid_facts.lint_facts(facts_list)
        # 占位符（「待补充」等）会让正文据此留空，属于硬拦截项；确需先写可 force 覆盖。
        placeholders = [item for item in issues if item.get("kind") == "placeholder"]
        empties = [item for item in issues if item.get("kind") == "empty_value"]
        if placeholders:
            names = "、".join(str(item.get("name")) for item in placeholders[:5])
            blockers.append(f"事实清单仍有占位符未落实：{names}（请先补真实数据，或 force 强行推进）")
        if empties:
            names = "、".join(str(item.get("name")) for item in empties[:5])
            warnings.append(f"事实值为空：{names}")
        # 建档时提交值曾与资质清单/重复值冲突（已按权威值修正），仅提示，不拦截。
        recorded_errors = [item for item in record.get("issues", []) if item.get("level") == "error"]
        if recorded_errors:
            warnings.append("建档时曾修正过冲突（已以资质清单为准）："
                            + "；".join(str(item.get("message")) for item in recorded_errors[:3]))

    elif name == "ledger_id":
        try:
            ledger = bid_consistency.load_ledger(str(value))
        except ValueError as exc:
            blockers.append(f"ledger_id 无效：{exc}")
            return blockers, warnings
        repair = ledger.get("repair")
        if not repair:
            blockers.append("一致性台账尚未给出统一口径：请先 cons-repair 提交 rules")
        elif not (repair.get("summary") or {}).get("ok"):
            unresolved = repair.get("unresolved") or []
            names = "、".join(str(item.get("name")) for item in unresolved[:5])
            blockers.append(f"一致性审计仍有未解决冲突：{names or '见台账'}")
        elif (ledger.get("summary") or {}).get("conflicts", 0) == 0:
            warnings.append("台账无冲突，一致性审计结论为「全部一致」")

    elif name == "files":
        names = value if isinstance(value, list) else [value]
        names = [_clip(item, 200) for item in names if _clip(item, 200)]
        if not names:
            blockers.append("files 为空：请先在 draft 生成 Word")
        for item in names:
            if item != Path(item).name:
                blockers.append(f"文件名非法：{item}")
            elif not (Path(settings.BID_OUTPUT_DIR) / item).is_file():
                blockers.append(f"Word 文件不存在：{item}")

    elif name == "rejection_session_id":
        try:
            state = bid_rejection.load_state(str(value))
        except ValueError as exc:
            blockers.append(f"rejection_session_id 无效：{exc}")
            return blockers, warnings
        if not (state.get("cursor") or {}).get("finished"):
            blockers.append("废标项检查尚未标记完成（cursor.finished=false），请跑完全部片段")
        pending = len(state.get("pending_risks") or [])
        if pending:
            warnings.append(f"废标项检查仍有 {pending} 条待定风险，须人工核对原文")

    return blockers, warnings


def _artifact_blockers(project: dict, index: int) -> tuple[list[str], list[str]]:
    """只看产物本身是否达标（不含「人工确认跳过」的豁免）。"""
    stage = STAGES[index]
    blockers: list[str] = []
    warnings: list[str] = []
    for name in stage["artifacts"]:
        stage_blockers, stage_warnings = _check_artifact(name, project.get("artifacts", {}).get(name))
        blockers.extend(f"[{stage['label']}] {item}" for item in stage_blockers)
        warnings.extend(f"[{stage['label']}] {item}" for item in stage_warnings)
    return blockers, warnings


def _artifact_fingerprint(name: str, value: Any) -> str:
    """产物的内容指纹：内容一变，之前的 force 确认就失效（避免过期豁免）。"""
    try:
        if name == "facts_id":
            record = bid_facts.load_facts(str(value))
            blob = json.dumps(record.get("facts", []), ensure_ascii=False, sort_keys=True)
            return hashlib.sha256(blob.encode("utf-8")).hexdigest()[:16]
        if name == "ledger_id":
            ledger = bid_consistency.load_ledger(str(value))
            blob = json.dumps({"repair": (ledger.get("repair") or {}).get("summary"),
                               "conflicts": (ledger.get("summary") or {}).get("conflicts"),
                               "unresolved": (ledger.get("repair") or {}).get("unresolved")},
                              ensure_ascii=False, sort_keys=True)
            return hashlib.sha256(blob.encode("utf-8")).hexdigest()[:16]
        if name == "rejection_session_id":
            state = bid_rejection.load_state(str(value))
            return "finished={}".format(bool((state.get("cursor") or {}).get("finished")))
        if name == "files":
            names = value if isinstance(value, list) else [value]
            return hashlib.sha256("|".join(sorted(str(item) for item in names)).encode("utf-8")).hexdigest()[:16]
    except ValueError:
        return "invalid"
    return "n/a"


def _signature(project: dict, index: int) -> str:
    stage = STAGES[index]
    values = {}
    for name in stage["artifacts"]:
        value = project.get("artifacts", {}).get(name)
        values[name] = {"id": value, "fp": _artifact_fingerprint(name, value)}
    return json.dumps(values, ensure_ascii=False, sort_keys=True)


def _accepted(project: dict, index: int) -> bool:
    stage = STAGES[index]
    entry = (project.get("accepted") or {}).get(stage["key"])
    return bool(entry) and entry.get("signature") == _signature(project, index)


def _stage_check(project: dict, index: int) -> tuple[list[str], list[str]]:
    """有效关卡结论：已 force 确认跳过的阻塞项降级为提示。"""
    blockers, warnings = _artifact_blockers(project, index)
    if blockers and _accepted(project, index):
        label = STAGES[index]["label"]
        warnings = [f"[{label}] 已人工确认跳过：{item}" for item in blockers] + warnings
        blockers = []
    return blockers, warnings


def _evaluate(project: dict, target_index: int) -> tuple[list[str], list[str]]:
    """校验「推进到 target_index 所代表的完成状态」是否成立：该阶段及其之前所有阶段的产物全部达标。"""
    blockers: list[str] = []
    warnings: list[str] = []
    for index in range(target_index + 1):
        stage_blockers, stage_warnings = _stage_check(project, index)
        blockers.extend(stage_blockers)
        warnings.extend(stage_warnings)
    return blockers, warnings


# ---------- 存取 ----------


def project_dir() -> Path:
    return Path(settings.bid_state_dir) / "project"


def _project_path(project_id: str) -> Path:
    if not _PROJECT_ID.fullmatch(str(project_id or "")):
        raise ProjectError(f"project_id 非法：{project_id!r}")
    return project_dir() / f"{project_id}.json"


def save_project(project: dict) -> Path:
    path = _project_path(str(project["project_id"]))
    path.parent.mkdir(parents=True, exist_ok=True)
    project["updated_at"] = _now()
    handle_fd, temporary = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    try:
        with os.fdopen(handle_fd, "w", encoding="utf-8") as handle:
            handle.write(json.dumps(project, ensure_ascii=False, indent=2))
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)
    return path


def load_project(project_id: str) -> dict:
    path = _project_path(project_id)
    if not path.is_file():
        raise ProjectError(f"未找到项目：{project_id}")
    try:
        project = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise ProjectError(f"项目无法解析：{type(exc).__name__}") from exc
    if not isinstance(project, dict) or int(project.get("version") or 0) != STATE_VERSION:
        raise ProjectError("项目版本不匹配，请重新初始化")
    return project


# ---------- 对外流程 ----------


def init_project(project: dict | None) -> dict:
    """创建一个项目工作流，落在第一个阶段。"""
    project = project if isinstance(project, dict) else {}
    project_id = f"prj-{time.strftime('%Y%m%d')}-{uuid.uuid4().hex[:8]}"
    record = {
        "version": STATE_VERSION,
        "project_id": project_id,
        "name": _clip(project.get("name") or project.get("project_name"), 120),
        "bid_no": _clip(project.get("bid_no"), 60),
        "created_at": _now(),
        "updated_at": _now(),
        "stage": STAGE_KEYS[0],
        "artifacts": {},
        "history": [{"at": _now(), "stage": STAGE_KEYS[0], "action": "init"}],
    }
    save_project(record)
    return project_state(project_id, record=record)


def project_state(project_id: str, *, record: dict | None = None) -> dict:
    """返回项目当前阶段、各阶段产物、以及当前阶段的关卡结论。"""
    project = record or load_project(project_id)
    current_index = _STAGE_INDEX.get(str(project.get("stage")), 0)
    blockers, warnings = _evaluate(project, current_index)
    stages = []
    for index, stage in enumerate(STAGES):
        if index > current_index:
            status = "pending"
        else:
            raw_blockers, _ = _artifact_blockers(project, index)
            if not raw_blockers:
                status = "done"
            elif _accepted(project, index):
                status = "accepted"      # 曾 force 跳过，产物仍未达标 —— 显示出来，别让它消失
            else:
                status = "blocked"
        stages.append({"key": stage["key"], "label": stage["label"], "status": status,
                       "artifacts": stage["artifacts"]})
    return {
        "project_id": project["project_id"],
        "name": project.get("name"),
        "bid_no": project.get("bid_no"),
        "stage": project.get("stage"),
        "stage_index": current_index,
        "stage_label": STAGES[current_index]["label"],
        "stages": stages,
        "artifacts": project.get("artifacts", {}),
        "blockers": blockers,
        "warnings": warnings,
        "can_advance": not blockers,
        "next": next_action(project),
        "history": project.get("history", []),
    }


def next_action(project: dict) -> dict:
    """给出「下一步该做什么」——直接可执行的指引，供 Agent 消费。"""
    current_index = _STAGE_INDEX.get(str(project.get("stage")), 0)
    stage = STAGES[current_index]
    blockers, _ = _evaluate(project, current_index)
    missing = [name for name in stage["artifacts"] if not project.get("artifacts", {}).get(name)]

    if blockers or missing:
        return {"action": "complete_stage", "stage": stage["key"], "label": stage["label"],
                "guide": stage["guide"], "blockers": blockers, "missing_artifacts": missing,
                "advance_with": stage["artifacts"]}

    if current_index + 1 >= len(STAGES):
        return {"action": "workflow_complete", "stage": stage["key"], "label": stage["label"],
                "guide": "全部阶段已完成。"}

    upcoming = STAGES[current_index + 1]
    return {"action": "advance", "stage": upcoming["key"], "label": upcoming["label"],
            "guide": upcoming["guide"], "advance_with": upcoming["artifacts"]}


def advance(project_id: str, stage: str, artifacts: dict | None = None,
            force: bool = False) -> dict:
    """声明「该阶段已完成」并推进到它。

    会严格校验该阶段及其之前所有阶段的产物是否真的达标；不达标则拒绝（除非 force），
    并逐条返回是哪一项不满足。`artifacts` 里不属于该阶段的字段会被拒绝（避免静默丢数据）。
    """
    project = load_project(project_id)
    stage_key = _clip(stage, 40)
    if stage_key not in _STAGE_INDEX:
        raise ProjectError(f"未知阶段：{stage_key!r}（可选：{'、'.join(STAGE_KEYS)}）")
    target_index = _STAGE_INDEX[stage_key]

    artifacts = artifacts if isinstance(artifacts, dict) else {}
    unknown = sorted(name for name in artifacts if name not in _ARTIFACT_NAMES)
    if unknown:
        raise ProjectError("不支持的产物字段：" + "、".join(unknown)
                           + f"（允许：{'、'.join(sorted(_ARTIFACT_NAMES))}）")

    for name in STAGES[target_index]["artifacts"]:
        if artifacts.get(name) not in (None, ""):
            project.setdefault("artifacts", {})[name] = artifacts[name]

    prior_blockers: list[str] = []
    warnings: list[str] = []
    for index in range(target_index):
        stage_blockers, stage_warnings = _stage_check(project, index)
        prior_blockers.extend(stage_blockers)
        warnings.extend(stage_warnings)
    target_blockers, target_warnings = _stage_check(project, target_index)
    warnings.extend(target_warnings)

    def _refuse(reason: str) -> dict:
        save_project(project)        # 保留已提交的产物，便于补齐后重试
        state = project_state(project_id, record=project)
        return {"ok": False, "stage": project.get("stage"), "target_stage": stage_key,
                "reason": reason, "blockers": prior_blockers + target_blockers,
                "warnings": warnings, "next": state["next"], "state": state}

    if prior_blockers:
        # force 只豁免「本阶段自身」的问题；前置阶段的问题要一级一级确认，
        # 否则一次 force 就能把整条链路的关卡全部踹开，关卡就失去意义了。
        return _refuse("前置阶段仍有未达标产物，请先修复（或在对应阶段逐级 force 确认）")
    if target_blockers and not force:
        return _refuse("本阶段产物未达标")

    forced = bool(target_blockers and force)
    if forced:
        project.setdefault("accepted", {})[stage_key] = {
            "signature": _signature(project, target_index),
            "blockers": target_blockers, "at": _now()}

    project["stage"] = stage_key
    project["history"].append({"at": _now(), "stage": stage_key,
                               "action": "force" if forced else "advance",
                               "forced_blockers": target_blockers if forced else []})
    save_project(project)
    state = project_state(project_id, record=project)
    return {"ok": True, "stage": stage_key, "forced": forced,
            "blockers": target_blockers, "warnings": warnings,
            "next": state["next"], "state": state}


def stages_doc() -> dict:
    """阶段清单（供文档与前端展示）。"""
    return {"stages": [{"key": stage["key"], "label": stage["label"],
                        "artifacts": stage["artifacts"], "guide": stage["guide"]} for stage in STAGES]}


__all__ = ["ProjectError", "STAGES", "STAGE_KEYS", "init_project", "project_state", "next_action",
           "advance", "load_project", "save_project", "project_dir", "stages_doc"]
