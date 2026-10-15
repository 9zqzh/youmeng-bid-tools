#!/usr/bin/env python3
"""优盟标书工具（技能本地版）。

**不需要任何服务**：本脚本直接调用随技能一起分发的实现（`scripts/bidagent/`），
在技能沙箱里就地跑。实现由后端仓库 `services/` 同步而来（见 `tools/sync_from_backend.py`）。

子命令与参数名与「thin client 版」保持一致，只有两点不同：
- 不再有下载链接 —— 产物直接写到**产物目录**（默认 `/workspace/output`），
  返回体里给的是 `path`；用户会在回复里看到可下载的文件；
- 出错时向 **stdout** 打印 `{"success": false, "error": "..."}` 并以退出码 1 结束，
  便于调用方统一解析。

依赖见 `requirements.txt`（python-docx / openpyxl / lxml / Pillow / pypdf / pypdfium2 / httpx，
配图渲染另需 playwright + chromium）。

环境变量（都有默认值，一般不用设）：
  BID_OUTPUT_DIR   产物目录           默认 /workspace/output
  BID_STATE_DIR    状态目录           默认 <技能目录>/.state（不可写则 /workspace/.bid-state）
  BID_INPUT_DIR    输入目录           默认 /workspace/input（放 manifest.json 与资质原件）
  BID_MANIFEST_PATH 资质清单路径      默认在输入目录里找 manifest.json
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# 让 `bidagent` 可被导入：脚本在 <技能>/scripts/ 下，运行时该目录即 sys.path[0]；
# 直接用 python 执行时通常已满足，这里再兜一层（远程执行的 cwd 可能不是脚本目录）。
sys.path.insert(0, str(Path(__file__).resolve().parent))

from bidagent import bid_api  # noqa: E402
from bidagent.settings import settings  # noqa: E402

# 会产出文件的子命令：返回体里补 output_dir，便于调用方告知用户去哪里取
ARTIFACT_COMMANDS = ("draft", "reject-export", "fig-render", "check-export")


def _load_json_arg(value: str):
    if value.startswith("@"):
        return json.loads(open(value[1:], encoding="utf-8").read())
    return json.loads(value)


def _load_text_arg(value: str) -> str:
    """HTML / 正文等纯文本参数：支持 @文件 或直接给字符串。"""
    if value.startswith("@"):
        return open(value[1:], encoding="utf-8").read()
    return value


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="bid_tools", description="优盟标书工具（技能本地版）")
    sub = parser.add_subparsers(dest="command", required=True)

    # ---------- 资质与初稿 ----------
    sub.add_parser("list", help="列出企业资质清单")

    p_match = sub.add_parser("match", help="资质逐条比对")
    p_match.add_argument("--requirements", required=True, help="要求数组 JSON 或 @file")

    p_draft = sub.add_parser("draft", help="生成两册 Word 初稿")
    p_draft.add_argument("--project", required=True)
    p_draft.add_argument("--sections", required=True, help="章节数组 JSON 或 @file")
    p_draft.add_argument("--qualifications", default="", help="要嵌入的资质 id，逗号分隔")
    p_draft.add_argument("--bid-no", default="")
    p_draft.add_argument("--final", action="store_true")
    p_draft.add_argument("--template", default="", help="导出模板名（预设 standard/a4-visual/a3-landscape 或自定义）")

    # ---------- 废标项检查 ----------
    p_ri = sub.add_parser("reject-init", help="创建废标项检查会话")
    p_ri.add_argument("--items", required=True, help="无效投标/废标项清单数组 JSON 或 @file")
    p_ri.add_argument("--project", default="", help="项目名称")
    p_ri.add_argument("--bid-no", default="")
    p_ri.add_argument("--documents", default="", help="待检查投标文件数组 JSON 或 @file")

    p_ra = sub.add_parser("reject-apply", help="提交一个片段的检查增量")
    p_ra.add_argument("--session", required=True, help="会话 id（reject-init 返回）")
    p_ra.add_argument("--patch", required=True, help="增量 patch 对象 JSON 或 @file")

    p_rf = sub.add_parser("reject-finalize", help="收尾并输出检查报告")
    p_rf.add_argument("--session", required=True)

    p_re = sub.add_parser("reject-export", help="导出废标项检查 xlsx")
    p_re.add_argument("--session", required=True)

    p_rs = sub.add_parser("reject-state", help="查看会话当前状态")
    p_rs.add_argument("--session", required=True)

    # ---------- 全局事实与一致性 ----------
    p_fi = sub.add_parser("facts-init", help="建立全局事实清单")
    p_fi.add_argument("--facts", required=True, help="事实数组 JSON 或 @file")
    p_fi.add_argument("--project", default="", help="项目名称")
    p_fi.add_argument("--bid-no", default="")
    p_fi.add_argument("--no-manifest", action="store_true", help="不注入资质清单的企业主体信息")

    p_fu = sub.add_parser("facts-update", help="按 id 全量替换事实清单")
    p_fu.add_argument("--facts-id", required=True)
    p_fu.add_argument("--facts", required=True, help="事实数组 JSON 或 @file")
    p_fu.add_argument("--no-manifest", action="store_true")

    p_fr = sub.add_parser("facts-render", help="取全局事实注入块（逐节生成正文时用）")
    p_fr.add_argument("--facts-id", required=True)

    p_cb = sub.add_parser("cons-build", help="一致性审计：建立台账并找出冲突")
    p_cb.add_argument("--sections", required=True, help="小节数组 JSON 或 @file")
    p_cb.add_argument("--extractions", required=True, help="逐节抽取结果数组 JSON 或 @file")
    p_cb.add_argument("--facts-id", default="", help="可选：关联全局事实以检出与全局冲突")

    p_cr = sub.add_parser("cons-repair", help="一致性审计：按统一口径生成逐节修复任务")
    p_cr.add_argument("--ledger", required=True, help="cons-build 返回的 ledger_id")
    p_cr.add_argument("--rules", required=True, help="统一口径数组 JSON 或 @file")

    # ---------- 项目工作流 ----------
    p_pi = sub.add_parser("project-init", help="创建标书项目（把全链路产物收拢到一个 project_id）")
    p_pi.add_argument("--project", default="", help="项目名称")
    p_pi.add_argument("--bid-no", default="")

    p_ps = sub.add_parser("project-state", help="查看项目：当前阶段 / 各阶段状态 / 关卡结论")
    p_ps.add_argument("--project-id", required=True)

    p_pn = sub.add_parser("project-next", help="下一步该做什么（给 Agent 直接消费）")
    p_pn.add_argument("--project-id", required=True)

    p_pa = sub.add_parser("project-advance", help="声明某阶段已完成并推进（带关卡校验）")
    p_pa.add_argument("--project-id", required=True)
    p_pa.add_argument("--stage", required=True,
                      help="analysis | facts | content | consistency | draft | rejection")
    p_pa.add_argument("--artifacts", default="", help="产物对象 JSON 或 @file，如 '{\"facts_id\":\"facts-..\"}'")
    p_pa.add_argument("--force", action="store_true", help="关卡不过也强行推进（会在项目里留痕）")

    # ---------- 配图 ----------
    sub.add_parser("fig-types", help="列出可渲染的配图类型与渲染方式（画图前先看）")
    sub.add_parser("fig-status", help="配图渲染环境自检")

    p_fig = sub.add_parser("fig-render", help="把 Mermaid/HTML 渲染成 PNG")
    p_fig.add_argument("--figures", required=True, help="图数组 JSON 或 @file")

    # ---------- 导出模板 ----------
    sub.add_parser("template-presets", help="列出导出模板预设")
    sub.add_parser("template-list", help="列出已保存的自定义模板")

    p_tg = sub.add_parser("template-get", help="把预设/模板名解析成完整配置")
    p_tg.add_argument("--template", default="", help="预设名或模板名；留空取 standard")

    p_ts = sub.add_parser("template-save", help="保存自定义导出模板")
    p_ts.add_argument("--name", required=True, help="模板名（字母数字与 _ -）")
    p_ts.add_argument("--template", required=True, help="模板对象 JSON 或 @file（可含 base 继承预设）")

    p_td = sub.add_parser("template-delete", help="删除自定义模板")
    p_td.add_argument("--name", required=True)

    # ---------- 受限 HTML ----------
    sub.add_parser("html-spec", help="受限 HTML 的白名单与示例（写 sections[].html 前先看）")
    p_hv = sub.add_parser("html-validate", help="校验一段受限 HTML")
    p_hv.add_argument("--html", required=True, help="HTML 字符串或 @文件")

    # ---------- 版面与配额 ----------
    p_lm = sub.add_parser("layout-metrics", help="算出一套模板下每页能装多少字")
    p_lm.add_argument("--template", default="", help="导出模板名（默认 standard）")

    p_lp = sub.add_parser("layout-plan", help="字数配额 + 版面估算 + 偏差与建议")
    p_lp.add_argument("--sections", required=True, help="小节数组 JSON 或 @file")
    p_lp.add_argument("--template", default="", help="导出模板名（默认 standard）")
    p_lp.add_argument("--total-words", type=int, default=0, help="总目标字数（按各节 weight 分配）")
    p_lp.add_argument("--section-words", type=int, default=0, help="每节统一目标字数")
    p_lp.add_argument("--min-words", type=int, default=0, help="单节目标下限")
    p_lp.add_argument("--max-words", type=int, default=0, help="单节目标上限")

    # ---------- 查重 ----------
    p_dup = sub.add_parser("dup-analyze", help="标书查重（确定性引擎：白名单剔除 + 分层比对）")
    p_dup.add_argument("--documents", required=True, help="待比对文件数组 JSON 或 @file")
    p_dup.add_argument("--tender", default="", help="招标原文（字符串或 @文件）；强烈建议提供，否则会大面积误报")
    p_dup.add_argument("--tender-path", default="", help="或给招标文件路径（本机可读即可）")

    # ---------- 正文检查 ----------
    p_tc = sub.add_parser("typo-check", help="回验错别字候选（定位不到的直接排除）")
    p_tc.add_argument("--candidates", required=True, help="候选数组 JSON 或 @file")
    p_tc.add_argument("--source", default="", help="待检查正文（字符串或 @文件）")
    p_tc.add_argument("--source-path", default="", help="或给本机正文文件路径")
    p_tc.add_argument("--tender", default="", help="招标原文（字符串或 @文件）；强烈建议提供")
    p_tc.add_argument("--tender-path", default="", help="或给本机招标文件路径")

    p_lc = sub.add_parser("logic-check", help="逻辑谬误检查（跨段前后矛盾）")
    p_lc.add_argument("--segments", required=True, help="小节数组 JSON 或 @file")
    p_lc.add_argument("--extractions", required=True, help="逐段抽取结果数组 JSON 或 @file")

    p_ce = sub.add_parser("check-export", help="把 废标项/错别字/逻辑 导出到同一份 xlsx")
    p_ce.add_argument("--rejection-session", default="", help="废标项检查 session_id（可选）")
    p_ce.add_argument("--typo-id", default="", help="typo-check 返回的 review_id（可选）")
    p_ce.add_argument("--logic-id", default="", help="logic-check 返回的 review_id（可选）")
    p_ce.add_argument("--project", default="", help="项目名称（仅正文检查时用于命名）")

    return parser


def dispatch(args) -> dict:
    """把命令行参数映射到 `bidagent.bid_api`（子命令与参数名与 thin client 版一致）。"""
    command = args.command

    if command == "list":
        return bid_api.list_qualifications_impl()
    if command == "match":
        return bid_api.qualify_match_impl(_load_json_arg(args.requirements))
    if command == "draft":
        return bid_api.generate_draft_impl(
            args.project, _load_json_arg(args.sections),
            [q.strip() for q in args.qualifications.split(",") if q.strip()],
            args.bid_no, args.final, args.template or None)

    if command == "reject-init":
        return bid_api.rejection_init_impl(
            {"name": args.project, "bid_no": args.bid_no},
            _load_json_arg(args.items),
            _load_json_arg(args.documents) if args.documents else [])
    if command == "reject-apply":
        return bid_api.rejection_apply_impl(args.session, _load_json_arg(args.patch))
    if command == "reject-finalize":
        return bid_api.rejection_finalize_impl(args.session)
    if command == "reject-export":
        return bid_api.rejection_export_impl(args.session)
    if command == "reject-state":
        return bid_api.rejection_state_impl(args.session)

    if command == "facts-init":
        return bid_api.facts_init_impl({"name": args.project, "bid_no": args.bid_no},
                                       _load_json_arg(args.facts), not args.no_manifest)
    if command == "facts-update":
        return bid_api.facts_update_impl(args.facts_id, _load_json_arg(args.facts), not args.no_manifest)
    if command == "facts-render":
        return bid_api.facts_render_impl(args.facts_id)
    if command == "cons-build":
        return bid_api.consistency_build_impl(
            {"sections": _load_json_arg(args.sections), "extractions": _load_json_arg(args.extractions),
             "facts_id": args.facts_id or None})
    if command == "cons-repair":
        return bid_api.consistency_repair_impl(args.ledger, _load_json_arg(args.rules))

    if command == "project-init":
        return bid_api.project_init_impl({"name": args.project, "bid_no": args.bid_no})
    if command == "project-state":
        return bid_api.project_state_impl(args.project_id)
    if command == "project-next":
        return bid_api.project_next_impl(args.project_id)
    if command == "project-advance":
        return bid_api.project_advance_impl(
            args.project_id, args.stage,
            _load_json_arg(args.artifacts) if args.artifacts else {}, args.force)

    if command == "fig-types":
        return bid_api.figure_types_impl()
    if command == "fig-status":
        return bid_api.figure_status_impl()
    if command == "fig-render":
        return bid_api.figure_render_impl(_load_json_arg(args.figures))

    if command == "template-presets":
        return bid_api.template_presets_impl()
    if command == "template-list":
        return bid_api.template_list_impl()
    if command == "template-get":
        return bid_api.template_get_impl(args.template or None)
    if command == "template-save":
        return bid_api.template_save_impl(args.name, _load_json_arg(args.template))
    if command == "template-delete":
        return bid_api.template_delete_impl(args.name)

    if command == "html-spec":
        return bid_api.html_spec_impl()
    if command == "html-validate":
        return bid_api.html_validate_impl(_load_text_arg(args.html))

    if command == "layout-metrics":
        return bid_api.layout_metrics_impl(args.template or None)
    if command == "layout-plan":
        payload = {"sections": _load_json_arg(args.sections), "template": args.template or None}
        for key, value in (("total_words", args.total_words), ("section_words", args.section_words),
                           ("min_words", args.min_words), ("max_words", args.max_words)):
            if value:
                payload[key] = value
        return bid_api.layout_plan_impl(payload)

    if command == "dup-analyze":
        return bid_api.duplicate_analyze_impl({
            "documents": _load_json_arg(args.documents),
            "tender_text": _load_text_arg(args.tender) if args.tender else None,
            "tender_path": args.tender_path or None})

    if command == "typo-check":
        return bid_api.typo_check_impl({
            "candidates": _load_json_arg(args.candidates),
            "source_text": _load_text_arg(args.source) if args.source else None,
            "source_path": args.source_path or None,
            "tender_text": _load_text_arg(args.tender) if args.tender else None,
            "tender_path": args.tender_path or None})
    if command == "logic-check":
        return bid_api.logic_check_impl({"segments": _load_json_arg(args.segments),
                                         "extractions": _load_json_arg(args.extractions)})
    if command == "check-export":
        return bid_api.check_export_impl({
            "rejection_session_id": args.rejection_session or None,
            "typo_id": args.typo_id or None, "logic_id": args.logic_id or None,
            "project_name": args.project})

    raise ValueError(f"未知子命令：{command}")


def _print_json(payload: dict) -> None:
    """打印 JSON。终端装不下中文字符时（如 Windows GBK 控制台）退回 ASCII 转义，内容不失真。"""
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    try:
        print(text)
    except UnicodeEncodeError:
        print(json.dumps(payload, ensure_ascii=True, indent=2))


def main() -> int:
    args = build_parser().parse_args()
    try:
        result = dispatch(args)
    except ValueError as exc:
        # 一律打到 stdout，便于调用方统一解析；同时以非零码结束
        _print_json({"success": False, "error": str(exc)})
        return 1
    except Exception as exc:  # noqa: BLE001 - 兜底，避免把栈直接甩给模型
        _print_json({"success": False, "error": f"{type(exc).__name__}: {exc}"[:400]})
        return 1

    if args.command in ARTIFACT_COMMANDS:
        result.setdefault("output_dir", settings.BID_OUTPUT_DIR)
    _print_json(result)
    return 0


if __name__ == "__main__":
    sys.exit(main())
