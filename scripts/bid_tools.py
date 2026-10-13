#!/usr/bin/env python3
"""优盟标书工具（Skill 版，HTTP 瘦客户端）。

不自带任何业务代码，只通过 HTTP 调用「优盟标书工具服务」（bid-mcp）的 REST 接口：
  GET  /api/qualifications
  POST /api/qualify-match
  POST /api/generate-draft

子命令：
  list                                           列出企业资质清单
  match  --requirements <JSON|@file>             招标要求 × 资质逐条比对
  draft  --project <名称> --sections <JSON|@file> [--qualifications q15,q06]
         [--bid-no X] [--final]                   生成两册 Word 初稿

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

    args = parser.parse_args()

    if args.command == "list":
        result = _request("GET", "/api/qualifications")
    elif args.command == "match":
        result = _request("POST", "/api/qualify-match",
                          {"requirements": _load_json_arg(args.requirements)})
    else:
        result = _request("POST", "/api/generate-draft", {
            "project_name": args.project,
            "sections": _load_json_arg(args.sections),
            "qualification_ids": [q.strip() for q in args.qualifications.split(",") if q.strip()],
            "bid_no": args.bid_no,
            "final": args.final,
        })
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
