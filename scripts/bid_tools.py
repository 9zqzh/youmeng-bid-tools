#!/usr/bin/env python3
"""优盟标书工具（Skill 版）：与 MCP 服务功能一致，复用后端仓库的 services 逻辑。

子命令：
  list                                           列出企业资质清单
  match  --requirements <JSON|@file>             招标要求 × 资质逐条比对
  draft  --project <名称> --sections <JSON|@file> [--qualifications q015,q006]
         [--bid-no X] [--final] [--out DIR]      生成两册 Word 初稿

输出：JSON 到 stdout（供 Agent 解析）；失败时打印简短错误而非 traceback。

依赖：后端仓库 requirements.txt（python-docx / pypdf / pypdfium2 / Pillow / httpx / pydantic-settings）。
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path


def _bootstrap():
    """定位后端仓库根与 skill 资源，导入后端逻辑（须在 import services 之前设好环境变量）。"""
    here = Path(__file__).resolve()
    repo = Path(os.environ.get("BID_REPO") or here.parents[4]).resolve()
    assets = here.parents[1] / "assets"
    if not (repo / "services" / "bid_service.py").is_file():
        raise SystemExit(f"未找到后端仓库代码：{repo}/services（可用环境变量 BID_REPO 指定）")
    sys.path.insert(0, str(repo))
    os.environ.setdefault("BID_MANIFEST_PATH", str(assets / "manifest.json"))
    os.environ.setdefault("BID_OUTPUT_DIR", str(repo / "data" / "bid" / "outputs"))
    from services import bid_service  # noqa: PLC0415
    from services.bidding_documents import build_bid_documents  # noqa: PLC0415
    return bid_service, build_bid_documents


def _load_json_arg(value: str):
    """接受内联 JSON 字符串或 @文件路径。"""
    if value.startswith("@"):
        return json.loads(Path(value[1:]).read_text(encoding="utf-8"))
    return json.loads(value)


def _fail(message: str, code: int = 1) -> int:
    print(json.dumps({"success": False, "error": message}, ensure_ascii=False))
    return code


def main() -> int:
    parser = argparse.ArgumentParser(prog="bid_tools", description="优盟标书工具")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("list", help="列出企业资质清单")

    p_match = sub.add_parser("match", help="资质逐条比对")
    p_match.add_argument("--requirements", required=True, help="要求数组 JSON 或 @file")

    p_draft = sub.add_parser("draft", help="生成两册 Word 初稿")
    p_draft.add_argument("--project", required=True, help="项目名称")
    p_draft.add_argument("--sections", required=True, help="章节数组 JSON 或 @file")
    p_draft.add_argument("--qualifications", default="", help="要嵌入的资质 id，逗号分隔")
    p_draft.add_argument("--bid-no", default="")
    p_draft.add_argument("--final", action="store_true")
    p_draft.add_argument("--out", default="", help="输出目录（默认后端 data/bid/outputs）")

    args = parser.parse_args()

    try:
        bid_service, build_bid_documents = _bootstrap()
    except SystemExit as exc:
        return _fail(str(exc))

    try:
        if args.command == "list":
            manifest = bid_service.load_manifest()
            print(json.dumps({
                "company": manifest["company"],
                "items": [{"id": i.get("id"), "name": i.get("name"), "category": i.get("category"),
                           "valid_to": i.get("valid_to"), "keywords": i.get("keywords", [])}
                          for i in manifest.get("items", [])],
            }, ensure_ascii=False, indent=2))
            return 0

        if args.command == "match":
            requirements = _load_json_arg(args.requirements)
            if not isinstance(requirements, list):
                return _fail("requirements 必须是数组")
            result = bid_service.qualify_match(requirements, bid_service.load_manifest())
            print(json.dumps(result, ensure_ascii=False, indent=2))
            return 0

        # draft
        sections = _load_json_arg(args.sections)
        if not isinstance(sections, list):
            return _fail("sections 必须是数组")
        manifest = bid_service.load_manifest()
        missing: list[str] = []
        ids = [q.strip() for q in args.qualifications.split(",") if q.strip()]
        attachments = bid_service._resolve_attachments(
            [{"qualification_id": q, "volume": "business_technical"} for q in ids],
            manifest, missing, manifest.get("base_dir"))
        chapters = sections or bid_service._seed_sections(manifest["company"], manifest.get("profile"))
        out_dir = Path(args.out) if args.out else Path(os.environ["BID_OUTPUT_DIR"])
        result = build_bid_documents(out_dir, {"name": args.project, "bid_no": args.bid_no},
                                     chapters, attachments, None, final=args.final, missing=missing)
        print(json.dumps({
            "company": manifest["company"],
            "files": [{"name": f["name"], "path": f["path"], "kind": f["kind"]} for f in result["files"]],
            "warnings": result["warnings"],
        }, ensure_ascii=False, indent=2))
        return 0
    except ValueError as exc:
        return _fail(str(exc)[:300])
    except Exception as exc:  # noqa: BLE001
        return _fail(f"{type(exc).__name__}: {exc}"[:300])


if __name__ == "__main__":
    sys.exit(main())
