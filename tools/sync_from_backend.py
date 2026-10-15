#!/usr/bin/env python3
"""把后端 `digital-hub-test/services/` 的实现同步进本技能，保持两边一致。

**唯一的源是后端仓库**：改逻辑请改 `services/`，再跑本脚本重新生成。
手改 `scripts/bidagent/` 下的副本会在下次同步时被覆盖（`settings.py` 例外，它是技能侧手写的）。

用法：
    python tools/sync_from_backend.py                 # 同步（默认自动找后端仓库）
    python tools/sync_from_backend.py --backend <路径>
    python tools/sync_from_backend.py --check         # 只比较不写入，确认两边一致

同步时会做两件事：
1. 拷贝模块文件；
2. 改写 import：`from config.settings import settings` → `from .settings import settings`，
   `from services import x` → `from . import x`。

并写一份 `.synced.json`（后端 commit + 各文件 sha256），供 `--check` 与溯源使用。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL_ROOT = HERE.parent
TARGET_DIR = SKILL_ROOT / "scripts" / "bidagent"
SYNCED_FILE = TARGET_DIR / ".synced.json"

# 需要同步的模块（技能侧额外手写的 settings.py 不在此列）
MODULES = [
    "bid_api.py", "bid_service.py", "bidding_documents.py",
    "bid_consistency.py", "bid_duplicate.py", "bid_facts.py", "bid_figures.py",
    "bid_html.py", "bid_layout.py", "bid_project.py", "bid_rejection.py",
    "bid_review.py", "bid_template.py",
]

# import 改写规则（顺序有意义）。注意 `from services import (...)` 有跨行括号写法，
# 单行正则匹配不到 —— 两种都要覆盖。
REWRITES = (
    (re.compile(r"^from config\.settings import settings", re.M), "from .settings import settings"),
    (re.compile(r"^from services import \(([^)]*)\)", re.M), r"from . import (\1)"),
    (re.compile(r"^from services import ([A-Za-z_][\w, ]*)$", re.M), r"from . import \1"),
    (re.compile(r"^from services\.([A-Za-z_]\w*) import", re.M), r"from .\1 import"),
)


def find_backend(explicit: str | None) -> Path:
    """定位后端仓库：参数 > 环境变量 > 常见同级路径 > **逐级向上的祖先目录**。"""
    candidates = []
    if explicit:
        candidates.append(Path(explicit))
    import os
    if os.environ.get("BID_BACKEND_DIR"):
        candidates.append(Path(os.environ["BID_BACKEND_DIR"]))
    for name in ("digital-hub-test", "digital-hub"):
        candidates.append(SKILL_ROOT.parent / name)
        candidates.append(SKILL_ROOT.parent.parent / name)
    # 第二份副本在 digital-hub-test/.qoder/skills/ 下，后端就是它的祖先目录
    for ancestor in list(SKILL_ROOT.parents)[:5]:
        candidates.append(ancestor)
    for candidate in candidates:
        if (candidate / "services" / "bid_api.py").is_file():
            return candidate.resolve()
    raise SystemExit(
        "找不到后端仓库（需含 services/bid_api.py）。用 --backend <路径> 或设 BID_BACKEND_DIR。")


def rewrite(text: str) -> str:
    for pattern, replacement in REWRITES:
        text = pattern.sub(replacement, text)
    return text


def digest(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def backend_commit(backend: Path) -> str:
    try:
        return subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=backend,
                              capture_output=True, text=True, check=True).stdout.strip()
    except (OSError, subprocess.SubprocessError):
        return "unknown"


def main() -> int:
    parser = argparse.ArgumentParser(description="从后端 services/ 同步实现到本技能")
    parser.add_argument("--backend", default=None, help="后端仓库根目录")
    parser.add_argument("--check", action="store_true", help="只检查是否一致，不写入")
    args = parser.parse_args()

    backend = find_backend(args.backend)
    source_dir = backend / "services"
    commit = backend_commit(backend)
    print(f"后端仓库：{backend}  (commit {commit})")

    rendered: dict[str, str] = {}
    for name in MODULES:
        source = source_dir / name
        if not source.is_file():
            raise SystemExit(f"后端缺少模块：{source}")
        rendered[name] = rewrite(source.read_text(encoding="utf-8"))

    manifest = {"backend_commit": commit, "files": {name: digest(text) for name, text in rendered.items()}}

    if args.check:
        if not SYNCED_FILE.is_file():
            print("未同步过（缺 .synced.json）—— 请先跑一次不带 --check 的同步。")
            return 1
        recorded = json.loads(SYNCED_FILE.read_text(encoding="utf-8"))
        problems = []
        if recorded.get("backend_commit") != commit:
            problems.append(f"后端 commit 变了：{recorded.get('backend_commit')} → {commit}")
        for name, text in rendered.items():
            if recorded.get("files", {}).get(name) != digest(text):
                problems.append(f"{name} 与后端不一致")
        for name, text in rendered.items():
            target = TARGET_DIR / name
            if not target.is_file() or target.read_text(encoding="utf-8") != text:
                problems.append(f"技能侧 {name} 与预期不符")
        if problems:
            print("不一致：")
            for item in problems:
                print("  -", item)
            print("\n跑 `python tools/sync_from_backend.py` 重新同步。")
            return 1
        print("一致 (OK)")
        return 0

    TARGET_DIR.mkdir(parents=True, exist_ok=True)
    for name, text in rendered.items():
        # newline="\n"：无论 Windows/Linux，生成物一律 LF，与后端一致（也避免 CRLF 噪音 diff）
        (TARGET_DIR / name).write_text(text, encoding="utf-8", newline="\n")
        print(f"  写出 {name} ({len(text)} 字节)")
    SYNCED_FILE.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"\n完成：{len(rendered)} 个模块 -> {TARGET_DIR.relative_to(SKILL_ROOT)}")
    print("（settings.py 不参与同步，属技能侧手写）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
