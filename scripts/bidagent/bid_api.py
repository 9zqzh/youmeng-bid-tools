"""标书工具的业务入口层：MCP 工具、REST 端点与本地 CLI 共用同一组函数。

这一层只做「参数归一化 + 调用 `services/*`」，**不含传输细节**：
不拼下载链接、不碰 HTTP。链接由传输层（`mcp-server/mcp_server.py`）自己加。

这样同一份逻辑有两种消费方式：
- 服务端：`mcp_server.py` 包装成 MCP 工具与 REST 端点（给 WeKnora Agent / Skill 沙箱）；
- 沙箱内：skill 直接 import 本模块在本地跑（不需要任何服务；产物直接落在工作区）。
"""
from __future__ import annotations

from pathlib import Path

from .settings import settings
from . import (
    bid_consistency,
    bid_duplicate,
    bid_facts,
    bid_figures,
    bid_html,
    bid_layout,
    bid_project,
    bid_rejection,
    bid_review,
    bid_service,
    bid_template,
)
from .bidding_documents import build_bid_documents


# ---------- 资质清单与初稿 ----------


def list_qualifications_impl() -> dict:
    manifest = bid_service.load_manifest()
    return {
        "company": manifest["company"],
        "items": [
            {"id": item.get("id"), "name": item.get("name"),
             "category": item.get("category"), "valid_to": item.get("valid_to"),
             "keywords": item.get("keywords", [])}
            for item in manifest.get("items", [])
        ],
    }


def qualify_match_impl(requirements: list[dict]) -> dict:
    if not isinstance(requirements, list):
        raise ValueError("requirements 必须是数组")
    return bid_service.qualify_match(requirements, bid_service.load_manifest())


def generate_draft_impl(project_name: str, sections: list[dict], qualification_ids: list[str] | None = None,
                        bid_no: str = "", final: bool = False, template=None) -> dict:
    """生成两册 Word 初稿。返回 `files: [{name, path}]`（调用方负责加下载链接）。"""
    if not isinstance(sections, list):
        raise ValueError("sections 必须是数组")
    manifest = bid_service.load_manifest()
    missing: list[str] = []
    attachments = bid_service._resolve_attachments(
        [{"qualification_id": q, "volume": "business_technical"} for q in (qualification_ids or [])],
        manifest, missing, manifest.get("base_dir"))
    chapters = sections or bid_service._seed_sections(manifest["company"], manifest.get("profile"))
    result = build_bid_documents(
        Path(settings.BID_OUTPUT_DIR), {"name": project_name, "bid_no": bid_no},
        chapters, attachments, None, final=bool(final), missing=missing, template=template)
    return {
        "company": manifest["company"],
        "files": [{"name": item["name"], "path": item["path"]} for item in result["files"]],
        "warnings": result["warnings"],
    }


# ---------- 废标项检查（程序维护权威状态，模型只提交增量 patch） ----------


def rejection_init_impl(project: dict | None, rejection_items: list[dict],
                        documents: list[dict] | None = None) -> dict:
    state = bid_rejection.init_session(project, rejection_items, documents)
    return {
        "session_id": state["session_id"],
        "project": state["project"],
        "checklist": [
            {"id": item["id"], "type": item["type"], "title": item["title"],
             "requirement": item["requirement"], "scope": item["scope"], "applicable": item["applicable"]}
            for item in state["checklist"]
        ],
        "documents": state["documents"],
        "stats": bid_rejection.state_stats(state),
        "next": ("按文档顺序逐片段调用 rejection_apply 提交增量 patch；"
                 "全部片段完成后调用 rejection_finalize，再调用 rejection_export 导出 xlsx。"),
    }


def rejection_apply_impl(session_id: str, patch: dict) -> dict:
    state = bid_rejection.load_state(session_id)
    diagnostics = bid_rejection.apply_patch(state, patch)
    bid_rejection.save_state(state)
    return {"session_id": session_id, **diagnostics}


def rejection_finalize_impl(session_id: str) -> dict:
    return bid_rejection.finalize(bid_rejection.load_state(session_id))


def rejection_export_impl(session_id: str) -> dict:
    result = bid_rejection.export_report(session_id)
    return {"name": result["name"], "path": result["path"]}


def rejection_state_impl(session_id: str) -> dict:
    return bid_rejection.load_state(session_id)


# ---------- 全局事实与一致性审计 ----------


def _manifest_optional() -> dict | None:
    """取资质清单用于注入锁定事实；清单缺失不应阻断事实建档。"""
    try:
        return bid_service.load_manifest()
    except Exception:  # noqa: BLE001
        return None


def facts_init_impl(project: dict | None, facts: list[dict], use_manifest: bool = True) -> dict:
    return bid_facts.init_facts(project, facts, _manifest_optional() if use_manifest else None)


def facts_update_impl(facts_id: str, facts: list[dict], use_manifest: bool = True) -> dict:
    return bid_facts.update_facts(facts_id, facts, _manifest_optional() if use_manifest else None)


def facts_render_impl(facts_id: str) -> dict:
    return bid_facts.get_rendered(facts_id)


def consistency_build_impl(payload: dict) -> dict:
    return bid_consistency.build_ledger(payload)


def consistency_repair_impl(ledger_id: str, rules: list[dict]) -> dict:
    return bid_consistency.repair_plan(ledger_id, rules)


# ---------- 项目工作流（串起全链路 + 关卡校验） ----------


def project_init_impl(project: dict | None) -> dict:
    return bid_project.init_project(project)


def project_state_impl(project_id: str) -> dict:
    return bid_project.project_state(project_id)


def project_next_impl(project_id: str) -> dict:
    return bid_project.next_action(bid_project.load_project(project_id))


def project_advance_impl(project_id: str, stage: str, artifacts: dict | None = None,
                         force: bool = False) -> dict:
    return bid_project.advance(project_id, stage, artifacts, force)


def project_stages_impl() -> dict:
    return bid_project.stages_doc()


# ---------- 配图渲染（Mermaid / HTML → PNG） ----------


def figure_types_impl() -> dict:
    return bid_figures.figure_types()


def figure_status_impl() -> dict:
    return bid_figures.renderer_status()


def figure_render_impl(figures: list[dict]) -> dict:
    """渲染配图。返回的 `rendered[].path` 是落盘绝对路径（调用方负责加下载链接）。"""
    return bid_figures.render_figures(figures)


# ---------- Word 导出模板 ----------


def template_presets_impl() -> dict:
    return bid_template.list_presets()


def template_list_impl() -> dict:
    return bid_template.list_templates()


def template_get_impl(ref) -> dict:
    resolved = bid_template.resolve_template(ref)
    return {"name": resolved["name"], "label": resolved["label"], "template": resolved}


def template_save_impl(name: str, template) -> dict:
    return bid_template.save_template(name, template)


def template_delete_impl(name: str) -> dict:
    return bid_template.delete_template(name)


# ---------- 受限 HTML 正文 ----------


def html_spec_impl() -> dict:
    return bid_html.html_spec()


def html_validate_impl(html) -> dict:
    report = bid_html.validate_html(html)
    if report["ok"]:
        report["message"] = "校验通过，可以放进 sections[].html 出稿。"
    else:
        report["message"] = f"有 {len(report['issues'])} 处问题，请按 message 修正后重试。"
    return report


# ---------- 版面估算与字数配额 ----------


def layout_metrics_impl(template=None) -> dict:
    resolved = bid_template.resolve_template(template)
    return {"template": resolved["name"], "label": resolved["label"],
            "metrics": bid_layout.page_metrics(resolved)}


def layout_plan_impl(payload: dict) -> dict:
    return bid_layout.plan(
        payload.get("sections"), template=payload.get("template"),
        total_words=payload.get("total_words"), section_words=payload.get("section_words"),
        min_words=payload.get("min_words"), max_words=payload.get("max_words"))


# ---------- 标书查重（确定性引擎） ----------


def duplicate_analyze_impl(payload: dict) -> dict:
    return bid_duplicate.analyze(payload)


# ---------- 正文检查（错别字 / 逻辑谬误），与废标项检查共用导出 ----------


def typo_check_impl(payload: dict) -> dict:
    return bid_review.run_typo_check(
        source_text=payload.get("source_text"), candidates=payload.get("candidates"),
        tender_text=payload.get("tender_text"), source_name=payload.get("source_name") or "",
        source_path=payload.get("source_path"), tender_path=payload.get("tender_path"))


def logic_check_impl(payload: dict) -> dict:
    return bid_review.run_logic_check(payload.get("segments"), payload.get("extractions"),
                                      payload.get("facts_id"))


def check_export_impl(payload: dict) -> dict:
    result = bid_review.export_combined(
        rejection_session_id=payload.get("rejection_session_id"), typo_id=payload.get("typo_id"),
        logic_id=payload.get("logic_id"), name=payload.get("name"),
        project_name=payload.get("project_name") or "")
    return {"name": result["name"], "path": result["path"], "sheets": result.get("sheets", [])}


__all__ = [name for name in dir() if name.endswith("_impl")]
