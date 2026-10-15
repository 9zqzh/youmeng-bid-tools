"""配图渲染：把模型产出的 Mermaid 源码 / 受限 HTML 渲染为 PNG，供 Word 与前端使用。

为什么不直接让模型出图片：本期内网没有生图模型，而标书里绝大多数图是
**结构图与图表**（流程图 / 思维导图 / 甘特图 / 架构图 / 柱状图…），
用 Mermaid 与 HTML+SVG **确定性渲染**比生图更可控、可复现、也更清晰。

- **Mermaid**：图元 + 连线类（思维导图 / 流程图 / 实体关系图 / 时序图…）。
- **HTML**：图表与版面图（甘特图 / 组织架构 / 泳道 / RACI / 风险矩阵 / 系统架构 /
  拓扑 / WBS / 鱼骨 / 柱状 / 折线 / 饼 / 雷达…）。模型给自包含的 HTML+CSS+SVG。
- **AI 生图**：本期不做（需生图模型），语义类型表里列为 unsupported。

渲染走 Chromium（Playwright）。浏览器来源优先级：
`BID_CHROMIUM_PATH` > `BID_FIGURE_BROWSER_CHANNEL`（如 msedge/chrome，用系统浏览器）
> Playwright 自带 chromium。不可用时抛出带**可照做补救提示**的错误。

安全：渲染在服务端进行，因此**一律切断网络**（拦截并中止所有请求），
且 `html` 类**拒绝** `<script>` / `<iframe>` / 内联事件 / `javascript:` ——
模型只需输出静态 HTML+CSS+SVG。
"""
from __future__ import annotations

import asyncio
import concurrent.futures
import re
from pathlib import Path
from typing import Any

from .settings import settings


class FigureError(ValueError):
    """配图模块的可预期错误（对应 HTTP 422）。"""


# (语义类型 key, 中文标签, 渲染方式)
_TYPE_TABLE: tuple[tuple[str, str, str], ...] = (
    ("mindmap", "思维导图", "mermaid"),
    ("flowchart", "流程图", "mermaid"),
    ("er", "实体关系图", "mermaid"),
    ("sequence", "时序图", "mermaid"),
    ("state", "状态图", "mermaid"),
    ("class", "类图", "mermaid"),
    ("gantt", "甘特图", "html"),
    ("network", "进度网络图", "html"),
    ("org", "组织架构图", "html"),
    ("swimlane", "泳道图", "html"),
    ("raci", "RACI 职责矩阵", "html"),
    ("risk_matrix", "风险矩阵", "html"),
    ("architecture", "系统架构图", "html"),
    ("topology", "拓扑图", "html"),
    ("wbs", "WBS 工作分解结构图", "html"),
    ("fishbone", "鱼骨图", "html"),
    ("bar", "柱状图", "html"),
    ("line", "折线图", "html"),
    ("pie", "饼图", "html"),
    ("area", "面积图", "html"),
    ("radar", "雷达图", "html"),
    ("scatter", "散点图", "html"),
    ("timeline", "时间轴", "html"),
    ("process", "流程/步骤图", "html"),
)

# 需要生图模型，本期不支持（列出来是为了让模型知道「这类图现在画不了」，别硬塞 html）
_UNSUPPORTED_TABLE: tuple[tuple[str, str], ...] = (
    ("photo", "实物照片"),
    ("scene", "场景效果图"),
    ("cutout", "抠图/白底产品图"),
    ("cutaway", "剖视图"),
    ("exploded", "爆炸分解图"),
    ("render3d", "三维渲染图"),
)

FIGURE_KINDS = ("mermaid", "html")
_TYPE_TO_KIND = {key: kind for key, _, kind in _TYPE_TABLE}
# 中文标签也接受（模型有时直接写「流程图」）
_LABEL_TO_KEY = {label: key for key, label, _ in _TYPE_TABLE}
_MERMAID_TYPES = tuple(key for key, _, kind in _TYPE_TABLE if kind == "mermaid")

_SAFE_NAME = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,80}\.png$")
_SAFE_ID = re.compile(r"[^A-Za-z0-9_-]+")

# html 类的基本防护：只接受静态 HTML+CSS+SVG
_FORBIDDEN_HTML = (
    (re.compile(r"<\s*script", re.I), "<script>"),
    (re.compile(r"<\s*iframe", re.I), "<iframe>"),
    (re.compile(r"<\s*object|<\s*embed", re.I), "<object>/<embed>"),
    (re.compile(r"javascript\s*:", re.I), "javascript: 伪协议"),
    (re.compile(r"\son[a-z]+\s*=", re.I), "内联事件属性（如 onclick=）"),
)

_WRAPPER = """<!DOCTYPE html><html><head><meta charset="utf-8"><style>
  html, body { margin: 0; padding: 0; background: #ffffff; }
  body { font-family: "Noto Sans CJK SC", "Source Han Sans SC", "Microsoft YaHei",
         "PingFang SC", "WenQuanYi Zen Hei", sans-serif; }
  #root { padding: 8px; }
</style></head><body><!--BODY--></body></html>"""

# 注意：用 replace 而不是 str.format —— CSS 里的花括号会被 format 当成占位符。
_WRAPPER_HTML = _WRAPPER.replace("<!--BODY-->", '<div id="root"></div>')
_WRAPPER_MERMAID = _WRAPPER.replace("<!--BODY-->", '<div id="root"><div id="c"></div></div>')

_LAUNCH_ARGS = ["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu",
                "--hide-scrollbars", "--force-color-profile=srgb"]


# ---------- 语义类型路由 ----------


def figure_types() -> dict:
    """语义类型 → 渲染方式对照表（供模型选择，也供前端展示）。"""
    return {
        "supported": [{"type": key, "label": label, "kind": kind, "renderer": "mermaid" if kind == "mermaid" else "html+svg"}
                      for key, label, kind in _TYPE_TABLE],
        "unsupported": [{"type": key, "label": label, "reason": "需生图模型，本期不支持"}
                        for key, label in _UNSUPPORTED_TABLE],
        "kinds": list(FIGURE_KINDS),
        "hint": ("kind=mermaid 时提供 mermaid 源码；kind=html 时提供自包含的静态 HTML+CSS+SVG"
                 "（不要用 <script>，渲染时网络已切断）。"),
    }


def route_kind(value: str) -> str | None:
    """把语义类型（key 或中文标签）映射为渲染方式；未知返回 None。"""
    text = str(value or "").strip()
    if text in FIGURE_KINDS:
        return text
    key = _LABEL_TO_KEY.get(text, text)
    return _TYPE_TO_KIND.get(key)


# ---------- 输入归一化 ----------


def _clip(value: Any, limit: int) -> str:
    return str(value or "").strip()[:limit]


def normalize_figures(raw: Any) -> list[dict]:
    """归一化待渲染的图清单。每项 {id?, type?/kind, title?, mermaid|html, width?, height?, scale?}。"""
    if raw is None:
        raise FigureError("figures 不能为空")
    if isinstance(raw, dict):
        raw = raw.get("figures") if isinstance(raw.get("figures"), list) else [raw]
    if not isinstance(raw, list):
        raise FigureError("figures 必须是数组")

    figures: list[dict] = []
    seen: set[str] = set()
    for index, item in enumerate(raw, 1):
        if not isinstance(item, dict):
            continue
        raw_type = _clip(item.get("type"), 40)
        kind = _clip(item.get("kind"), 20) or route_kind(raw_type) or ""
        if kind not in FIGURE_KINDS:
            if raw_type in dict(_UNSUPPORTED_TABLE):
                label = dict(_UNSUPPORTED_TABLE)[raw_type]
                raise FigureError(
                    f"第 {index} 张图：类型「{label}」不在支持范围（需生图模型）。"
                    "支持的类型见 fig-types：结构图用 mermaid，图表/版面图用 html。")
            raise FigureError(
                f"第 {index} 张图：无法确定渲染方式，请显式给 kind=mermaid|html，或给受支持的 type。")

        source = _clip(item.get("mermaid") or item.get("code") or item.get("source"), 20000) \
            if kind == "mermaid" else _clip(item.get("html") or item.get("source"), 20000)
        if not source:
            field = "mermaid" if kind == "mermaid" else "html"
            raise FigureError(f"第 {index} 张图：kind={kind} 但缺少 {field} 内容")
        if kind == "html":
            for pattern, label in _FORBIDDEN_HTML:
                if pattern.search(source):
                    raise FigureError(
                        f"第 {index} 张图：html 含不被允许的内容（{label}）。"
                        "请只输出静态 HTML+CSS+SVG，渲染时网络已切断、脚本不会执行。")

        figure_id = _SAFE_ID.sub("", _clip(item.get("id"), 40)) or f"f{index}"
        if figure_id in seen:
            raise FigureError(f"第 {index} 张图：id 重复（{figure_id}）")
        seen.add(figure_id)

        name = _clip(item.get("name"), 100)
        if name and not _SAFE_NAME.fullmatch(name):
            raise FigureError(f"第 {index} 张图：name 非法（只允许字母数字与 . _ -，且需 .png 结尾）")

        figures.append({
            "id": figure_id,
            "type": raw_type or kind,
            "kind": kind,
            "title": _clip(item.get("title"), 120),
            "source": source,
            "width": max(320, min(int(item.get("width") or settings.BID_FIGURE_WIDTH), 4000)),
            "height": max(0, min(int(item.get("height") or 0), 6000)),
            "scale": max(1, min(int(item.get("scale") or settings.BID_FIGURE_SCALE), 4)),
            "name": name,
        })
    if not figures:
        raise FigureError("没有可渲染的图（figures 为空或格式不对）")
    return figures


# ---------- 浏览器与 mermaid 资源 ----------


def _mermaid_js() -> Path | None:
    candidates = []
    configured = (settings.BID_MERMAID_JS or "").strip()
    if configured:
        candidates.append(Path(configured))
    candidates.append(Path(__file__).resolve().parent.parent / "assets" / "mermaid.min.js")
    candidates.append(Path("/app/assets/mermaid.min.js"))
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


def renderer_status() -> dict:
    """渲染环境自检（不启动浏览器，故很快）。"""
    status: dict = {"playwright": False, "browser_configured": None,
                    "mermaid_js": None, "ready": False, "notes": []}
    try:
        import playwright  # noqa: F401
        status["playwright"] = True
    except ImportError:
        status["notes"].append("未安装 playwright（pip install playwright）")
    if (settings.BID_CHROMIUM_PATH or "").strip():
        status["browser_configured"] = f"BID_CHROMIUM_PATH={settings.BID_CHROMIUM_PATH}"
    elif (settings.BID_FIGURE_BROWSER_CHANNEL or "").strip():
        status["browser_configured"] = f"BID_FIGURE_BROWSER_CHANNEL={settings.BID_FIGURE_BROWSER_CHANNEL}"
    else:
        status["browser_configured"] = "Playwright 自带 chromium（未显式配置）"
        status["notes"].append("若容器内未执行 `playwright install chromium`，请改配 BID_CHROMIUM_PATH 或在镜像里安装")
    js = _mermaid_js()
    status["mermaid_js"] = str(js) if js else None
    if js is None:
        status["notes"].append("未找到 mermaid.min.js：kind=mermaid 的图会失败，html 类不受影响")
    status["ready"] = bool(status["playwright"])
    return status


def _launch(pw):
    """按优先级尝试启动浏览器，全部失败则抛出带补救提示的错误。"""
    attempts: list[tuple[str, dict]] = []
    if (settings.BID_CHROMIUM_PATH or "").strip():
        attempts.append(("BID_CHROMIUM_PATH", {"executable_path": settings.BID_CHROMIUM_PATH.strip()}))
    if (settings.BID_FIGURE_BROWSER_CHANNEL or "").strip():
        attempts.append(("BID_FIGURE_BROWSER_CHANNEL",
                         {"channel": settings.BID_FIGURE_BROWSER_CHANNEL.strip()}))
    attempts.append(("Playwright 自带 chromium", {}))

    errors: list[str] = []
    for label, kwargs in attempts:
        try:
            return pw.chromium.launch(args=_LAUNCH_ARGS, **kwargs), label
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{label}：{type(exc).__name__}: {str(exc)[:120]}")
    raise FigureError(
        "无法启动浏览器，配图渲染不可用。已尝试 → " + "；".join(errors)
        + "。补救：容器内执行 `playwright install --with-deps chromium`，"
          "或用 BID_CHROMIUM_PATH 指向系统 chromium，或 BID_FIGURE_BROWSER_CHANNEL=chrome。")


# ---------- 渲染 ----------


def figure_dir() -> Path:
    return Path(settings.BID_OUTPUT_DIR) / "figures"


def _render_one(browser, figure: dict, directory: Path) -> dict:
    mermaid_path: Path | None = None
    if figure["kind"] == "mermaid":
        mermaid_path = _mermaid_js()
        if mermaid_path is None:
            raise FigureError(
                "缺少 mermaid.min.js，无法渲染 mermaid 图。请下载后放到 <仓库>/assets/mermaid.min.js "
                "或用 BID_MERMAID_JS 指向文件（如 https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js）。"
                "也可把这些图改用 kind=html 由模型直接画。")

    name = figure["name"] or f"{figure['id']}.png"
    target_path = directory / name
    timeout_ms = max(5, int(settings.BID_FIGURE_TIMEOUT)) * 1000

    context = browser.new_context(
        viewport={"width": figure["width"], "height": figure["height"] or 900},
        device_scale_factor=figure["scale"])
    try:
        # 渲染在服务端进行，一律断网：模型产出的内容不应发起任何外部请求。
        context.route("**", lambda route: route.abort())
        page = context.new_page()
        page.set_default_timeout(timeout_ms)
        if figure["kind"] == "mermaid":
            page.set_content(_WRAPPER_MERMAID)
            page.add_script_tag(path=str(mermaid_path))
            page.evaluate(
                """async (code) => {
                     mermaid.initialize({startOnLoad: false, theme: 'neutral', securityLevel: 'strict'});
                     const r = await mermaid.render('fig', code);
                     document.getElementById('c').innerHTML = r.svg;
                     return true;
                   }""", figure["source"])
            target = page.locator("#c")
        else:
            page.set_content(_WRAPPER_HTML)
            page.evaluate("(html) => { document.getElementById('root').innerHTML = html; }",
                          figure["source"])
            target = page.locator("#root")
        target.screenshot(path=str(target_path))
    except FigureError:
        raise
    except Exception as exc:  # noqa: BLE001
        message = str(exc).splitlines()[0][:260] if str(exc) else type(exc).__name__
        raise FigureError(f"渲染失败：{message}") from exc
    finally:
        context.close()

    from PIL import Image
    with Image.open(target_path) as image:
        width, height = image.size
    return {"id": figure["id"], "type": figure["type"], "kind": figure["kind"],
            "title": figure["title"], "name": name, "path": str(target_path.resolve()),
            "width": width, "height": height, "bytes": target_path.stat().st_size}


def _run_offloaded(function, *args):
    """在需要时把渲染丢到工作线程执行。

    Playwright 的**同步 API 不允许在 asyncio 事件循环里运行**（会直接抛
    "It looks like you are using Playwright Sync API inside the asyncio loop"）。
    而本服务是被 FastAPI/Starlette 的异步路由调用的，所以这里检测到事件循环时
    自动切到工作线程；纯同步调用（如单测）直接执行。
    """
    try:
        asyncio.get_running_loop()
    except RuntimeError:
        return function(*args)
    with concurrent.futures.ThreadPoolExecutor(max_workers=1,
                                               thread_name_prefix="bid-figure") as pool:
        return pool.submit(function, *args).result()


def _render_all(figures: list[dict], directory: Path) -> dict:
    from playwright.sync_api import sync_playwright

    rendered: list[dict] = []
    failed: list[dict] = []
    with sync_playwright() as pw:
        browser, _label = _launch(pw)
        try:
            for figure in figures:
                try:
                    rendered.append(_render_one(browser, figure, directory))
                except FigureError as exc:
                    failed.append({"id": figure["id"], "type": figure["type"], "error": str(exc)[:300]})
                except Exception as exc:  # noqa: BLE001
                    failed.append({"id": figure["id"], "type": figure["type"],
                                   "error": f"{type(exc).__name__}: {str(exc)[:200]}"})
        finally:
            try:
                browser.close()
            except Exception:  # noqa: BLE001
                pass
    return {"rendered": rendered, "failed": failed,
            "summary": {"total": len(figures), "rendered": len(rendered), "failed": len(failed),
                        "ok": not failed}}


def render_figures(raw: Any) -> dict:
    """批量渲染。逐张容错：单张失败不影响其它，失败原因放在 failed 里。"""
    figures = normalize_figures(raw)
    directory = figure_dir()
    directory.mkdir(parents=True, exist_ok=True)

    try:
        import playwright  # noqa: F401
    except ImportError as exc:
        raise FigureError(
            "未安装 playwright，配图渲染不可用。请 `pip install playwright && playwright install --with-deps chromium`。"
        ) from exc

    try:
        return _run_offloaded(_render_all, figures, directory)
    except FigureError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise FigureError(f"渲染失败：{type(exc).__name__}: {str(exc)[:200]}") from exc


__all__ = ["FigureError", "FIGURE_KINDS", "figure_types", "route_kind", "normalize_figures",
           "render_figures", "renderer_status", "figure_dir"]
