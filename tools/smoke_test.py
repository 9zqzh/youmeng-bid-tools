#!/usr/bin/env python3
"""技能自检：确认依赖装好、且各子命令在本机可用（**不需要任何服务**）。

技能安装之后（沙箱里也一样）跑一次，就能知道「装没装好、能不能用」：

    python tools/smoke_test.py

它会先查依赖导入，再逐项调用子命令。**缺少资质清单时**与清单相关的项会被标成 SKIP
（不是 FAIL）—— 那说明还没把 `manifest.json` 放进输入目录。
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parent.parent
CLI = SKILL_ROOT / "scripts" / "bid_tools.py"

# 终端编码装不下中文时（如 Windows GBK 控制台）用替换字符兜底，别让自检自己崩掉
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass

results: list[tuple[str, str, str]] = []   # (状态, 项目, 说明)


def record(status: str, name: str, detail: str = "") -> None:
    results.append((status, name, detail))
    mark = {"PASS": "  OK  ", "FAIL": " FAIL ", "SKIP": " SKIP "}[status]
    print(f"{mark} {name}" + (f"  <- {detail}" if detail else ""))


def run(command: str, *args: str, expect_ok: bool = True) -> tuple[bool, dict | str]:
    """调用 CLI，返回 (是否成功, 解析后的 JSON 或原始文本)。"""
    proc = subprocess.run([sys.executable, str(CLI), command, *args],
                          capture_output=True, text=True, encoding="utf-8",
                          env={**os.environ, "PYTHONIOENCODING": "utf-8"})
    output = (proc.stdout or "").strip()
    try:
        payload = json.loads(output)
    except ValueError:
        payload = output or (proc.stderr or "").strip()
    ok = proc.returncode == 0 and not (isinstance(payload, dict) and payload.get("success") is False)
    if expect_ok and not ok:
        record("FAIL", f"{command}", str(payload)[:200])
    return ok, payload


def check_dependencies() -> bool:
    missing = []
    for module, package in (("docx", "python-docx"), ("openpyxl", "openpyxl"), ("lxml", "lxml"),
                            ("PIL", "Pillow"), ("pypdf", "pypdf"), ("pypdfium2", "pypdfium2"),
                            ("httpx", "httpx")):
        try:
            __import__(module)
        except ImportError:
            missing.append(package)
    if missing:
        record("FAIL", "依赖导入", "缺：" + "、".join(missing) + "（先按 requirements.txt 装）")
        return False
    record("PASS", "依赖导入", "python-docx / openpyxl / lxml / Pillow / pypdf / pypdfium2 / httpx")
    try:
        import playwright  # noqa: F401
        record("PASS", "playwright", "已装（配图渲染还需 chromium，见 fig-status）")
    except ImportError:
        record("SKIP", "playwright", "未装 —— 仅影响配图渲染，其它功能不受影响")
    return True


def main() -> int:
    print(f"技能目录：{SKILL_ROOT}")
    print(f"产物目录：{os.environ.get('BID_OUTPUT_DIR') or '/workspace/output'}")
    print(f"输入目录：{os.environ.get('BID_INPUT_DIR') or '/workspace/input'}")
    print()

    if not check_dependencies():
        return 1

    # ---- 不依赖资质清单的项 ----
    ok, presets = run("template-presets")
    if ok and isinstance(presets, dict):
        record("PASS", "template-presets", f"{len(presets.get('presets', []))} 个预设")
    ok, spec = run("html-spec")
    if ok and isinstance(spec, dict):
        record("PASS", "html-spec", f"{len(spec.get('block_tags', []))} 个块级标签")
    ok, metrics = run("layout-metrics")
    if ok and isinstance(metrics, dict):
        record("PASS", "layout-metrics", f"每页约 {metrics['metrics']['chars_per_page']} 字")
    # 失败也不算错：没 chromium 时它会给出补救提示，属正常
    status, _ = run("fig-status", expect_ok=False)
    record("PASS" if status else "SKIP", "fig-status", "配图渲染环境自检（未装 chromium 时属正常）")
    run("project-init", "--project", "自检", expect_ok=False)

    # ---- 需要资质清单的项 ----
    ok, quals = run("list", expect_ok=False)
    if not ok:
        record("SKIP", "list", "读不到资质清单 —— 把 manifest.json 放进输入目录再跑一次")
        print()
        print("结果：依赖与无清单项已通过；清单相关项跳过。")
        return 0
    items = quals.get("items", [])
    record("PASS", "list", f"{len(items)} 条资质")

    with tempfile.TemporaryDirectory():
        # 废标项：建档 → 提增量 → 收尾 → 导出
        sid = None
        ok, session = run("reject-init", "--project", "自检",
                          "--items", '[{"title":"未按格式提交投标函","requirement":"须按第六章格式"}]',
                          "--documents", '[{"name":"投标文件.pdf"}]')
        if ok:
            sid = session["session_id"]
            with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as handle:
                json.dump({"document_id": "doc_0001", "chunk_index": 0, "finished": True,
                           "confirmed_risk_adds": [{"checklist_id": "ri_0001", "severity": "high",
                                                    "title": "格式不符", "requirement": "须按格式",
                                                    "bid_evidence": "第 2 页抬头不符"}]}, handle)
                patch_file = handle.name
            run("reject-apply", "--session", sid, "--patch", f"@{patch_file}")
            run("reject-finalize", "--session", sid)
            ok, exported = run("reject-export", "--session", sid)
            if ok and Path(exported["path"]).is_file():
                record("PASS", "reject-export", f"xlsx 已生成：{exported['name']}")
            else:
                record("FAIL", "reject-export", "xlsx 未落盘")

        # 生成 Word（不带资质原件，验证链路本身）
        sections = json.dumps([{"title": "投标函", "content": "致：（招标人）",
                               "volume": "business_technical"},
                              {"title": "报价说明", "content": "见附表", "volume": "price"}],
                              ensure_ascii=False)
        ok, drafted = run("draft", "--project", "自检", "--sections", sections)
        if ok:
            files = drafted.get("files", [])
            missing = [f["name"] for f in files if not Path(f.get("path", "")).is_file()]
            if files and not missing:
                record("PASS", "draft", "两册 Word 已生成：" + "、".join(f["name"] for f in files))
            else:
                record("FAIL", "draft", f"文件未落盘：{missing}")
        # 合并导出（拿刚做出来的废标项检查作输入）
        if sid:
            ok, combined = run("check-export", "--project", "自检", "--rejection-session", sid)
            if ok:
                record("PASS", "check-export", "表：" + "/".join(combined.get("sheets") or ["（无附加表）"]))
        else:
            record("SKIP", "check-export", "没有可导出的检查结果")

    print()
    failed = [name for status, name, _ in results if status == "FAIL"]
    if failed:
        print("未通过：" + "、".join(failed))
        return 1
    print("自检通过。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
