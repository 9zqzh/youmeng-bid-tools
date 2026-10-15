#!/usr/bin/env python3
"""优盟标书工具（Skill 版，HTTP 瘦客户端）。

不自带任何业务代码，只通过 HTTP 调用「优盟标书工具服务」（bid-mcp）的 REST 接口：
  GET  /api/qualifications
  POST /api/qualify-match
  POST /api/generate-draft
  POST /api/rejection/init | /apply | /finalize | /export     GET /api/rejection/state
  POST /api/facts/init | /update | /render
  POST /api/consistency/build | /repair
  POST /api/project/init | /advance        GET /api/project/state | /next | /stages
  GET  /api/figure/types | /status         POST /api/figure/render

子命令：
  list                                           列出企业资质清单
  match  --requirements <JSON|@file>             招标要求 × 资质逐条比对
  draft  --project <名称> --sections <JSON|@file> [--qualifications q15,q06]
         [--bid-no X] [--final]                   生成两册 Word 初稿
  reject-init   --items <JSON|@file> [--project X] [--bid-no X] [--documents <JSON|@file>]
                                                 创建废标项检查会话
  reject-apply  --session <id> --patch <JSON|@file>
                                                 提交一个片段的检查增量
  reject-finalize --session <id>                 收尾并输出检查报告
  reject-export --session <id>                   导出废标项检查 xlsx
  reject-state  --session <id>                   查看会话当前状态
  facts-init    --facts <JSON|@file> [--project X] [--bid-no X] [--no-manifest]
                                                 建立「全局事实」清单（注入资质清单锁定事实）
  facts-update  --facts-id <id> --facts <JSON|@file> [--no-manifest]
                                                 人工评审后全量替换事实清单
  facts-render  --facts-id <id>                  取事实注入块（逐节生成正文时用）
  cons-build    --sections <JSON|@file> --extractions <JSON|@file> [--facts-id X]
                                                 一致性审计：建立台账并找出冲突
  cons-repair   --ledger <id> --rules <JSON|@file>
                                                 一致性审计：按统一口径生成逐节修复任务
  project-init  [--project X] [--bid-no X]        创建标书项目（串起全链路）
  project-state --project-id <id>                查看项目阶段与关卡结论
  project-next  --project-id <id>                下一步该做什么
  project-advance --project-id <id> --stage <key> [--artifacts <JSON|@file>] [--force]
                                                 声明某阶段完成并推进（带关卡校验）
  fig-types                                      列出可渲染的配图类型与渲染方式
  fig-status                                     配图渲染环境自检
  fig-render    --figures <JSON|@file>           把 Mermaid/HTML 渲染成 PNG（返回下载链接）

仅用 Python 标准库（urllib），无需 pip 安装任何依赖。

环境变量：
  BID_API_BASE   服务地址，默认 http://bid-mcp:18999
                 （跨机访问时改成 http://<服务器IP>:18999）
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import urllib.error
import urllib.request

BASE = (os.environ.get("BID_API_BASE") or "http://bid-mcp:18999").rstrip("/")
TIMEOUT = int(os.environ.get("BID_API_TIMEOUT", "180"))


def _request(method: str, path: str, payload: dict | None = None):
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(BASE + path, data=data, method=method,
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        try:
            detail = json.loads(exc.read().decode("utf-8"))
        except Exception:  # noqa: BLE001
            detail = {"error": f"HTTP {exc.code}"}
        raise SystemExit(json.dumps({"success": False, **detail}, ensure_ascii=False))
    except urllib.error.URLError as exc:
        raise SystemExit(json.dumps(
            {"success": False, "error": f"无法连接 {BASE}：{exc.reason}（用 BID_API_BASE 指定服务地址）"},
            ensure_ascii=False))


def _load_json_arg(value: str):
    if value.startswith("@"):
        return json.loads(open(value[1:], encoding="utf-8").read())
    return json.loads(value)


def main() -> int:
    parser = argparse.ArgumentParser(prog="bid_tools", description="优盟标书工具（HTTP 客户端）")
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("list", help="列出企业资质清单")

    p_match = sub.add_parser("match", help="资质逐条比对")
    p_match.add_argument("--requirements", required=True, help="要求数组 JSON 或 @file")

    p_draft = sub.add_parser("draft", help="生成两册 Word 初稿")
    p_draft.add_argument("--project", required=True)
    p_draft.add_argument("--sections", required=True, help="章节数组 JSON 或 @file")
    p_draft.add_argument("--qualifications", default="", help="要嵌入的资质 id，逗号分隔")
    p_draft.add_argument("--bid-no", default="")
    p_draft.add_argument("--final", action="store_true")

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

    sub.add_parser("fig-types", help="列出可渲染的配图类型与渲染方式（画图前先看）")
    sub.add_parser("fig-status", help="配图渲染环境自检")

    p_fr = sub.add_parser("fig-render", help="把 Mermaid/HTML 渲染成 PNG")
    p_fr.add_argument("--figures", required=True, help="图数组 JSON 或 @file")

    args = parser.parse_args()

    if args.command == "list":
        result = _request("GET", "/api/qualifications")
    elif args.command == "match":
        result = _request("POST", "/api/qualify-match",
                          {"requirements": _load_json_arg(args.requirements)})
    elif args.command == "draft":
        result = _request("POST", "/api/generate-draft", {
            "project_name": args.project,
            "sections": _load_json_arg(args.sections),
            "qualification_ids": [q.strip() for q in args.qualifications.split(",") if q.strip()],
            "bid_no": args.bid_no,
            "final": args.final,
        })
    elif args.command == "reject-init":
        result = _request("POST", "/api/rejection/init", {
            "project": {"name": args.project, "bid_no": args.bid_no},
            "rejection_items": _load_json_arg(args.items),
            "documents": _load_json_arg(args.documents) if args.documents else [],
        })
    elif args.command == "reject-apply":
        result = _request("POST", "/api/rejection/apply", {
            "session_id": args.session,
            "patch": _load_json_arg(args.patch),
        })
    elif args.command == "reject-finalize":
        result = _request("POST", "/api/rejection/finalize", {"session_id": args.session})
    elif args.command == "reject-export":
        result = _request("POST", "/api/rejection/export", {"session_id": args.session})
    elif args.command == "reject-state":
        result = _request("GET", f"/api/rejection/state?session_id={args.session}")
    elif args.command == "facts-init":
        result = _request("POST", "/api/facts/init", {
            "project": {"name": args.project, "bid_no": args.bid_no},
            "facts": _load_json_arg(args.facts),
            "use_manifest": not args.no_manifest,
        })
    elif args.command == "facts-update":
        result = _request("POST", "/api/facts/update", {
            "facts_id": args.facts_id,
            "facts": _load_json_arg(args.facts),
            "use_manifest": not args.no_manifest,
        })
    elif args.command == "facts-render":
        result = _request("POST", "/api/facts/render", {"facts_id": args.facts_id})
    elif args.command == "cons-build":
        result = _request("POST", "/api/consistency/build", {
            "facts_id": args.facts_id,
            "sections": _load_json_arg(args.sections),
            "extractions": _load_json_arg(args.extractions),
        })
    elif args.command == "cons-repair":
        result = _request("POST", "/api/consistency/repair", {
            "ledger_id": args.ledger,
            "rules": _load_json_arg(args.rules),
        })
    elif args.command == "project-init":
        result = _request("POST", "/api/project/init", {
            "project": {"name": args.project, "bid_no": args.bid_no}})
    elif args.command == "project-state":
        result = _request("GET", f"/api/project/state?project_id={args.project_id}")
    elif args.command == "project-next":
        result = _request("GET", f"/api/project/next?project_id={args.project_id}")
    elif args.command == "fig-types":
        result = _request("GET", "/api/figure/types")
    elif args.command == "fig-status":
        result = _request("GET", "/api/figure/status")
    elif args.command == "fig-render":
        result = _request("POST", "/api/figure/render", {"figures": _load_json_arg(args.figures)})
    else:
        result = _request("POST", "/api/project/advance", {
            "project_id": args.project_id,
            "stage": args.stage,
            "artifacts": _load_json_arg(args.artifacts) if args.artifacts else {},
            "force": args.force,
        })
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
