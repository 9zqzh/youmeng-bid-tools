"""受限 HTML 正文：让模型用受控 HTML 表达正文结构，程序校验后确定性转成 Word。

为什么不用纯文本行：`sections[].content` 是「按换行拆段落」，表达不了列表、子标题、
表格与图注 —— 模型只能靠空行和手写编号凑结构，转 Word 时结构信息全丢。

受限 HTML 的价值在于**可被程序机器校验**：标签/属性走白名单，越界就精确报错（而不是
静默删掉），模型能一次改对；再由程序确定性转 Word。这里**不经 LibreOffice 往返**，
而是直接映射到 python-docx 元素，避免格式不确定与额外服务依赖。

白名单：

    块级   p  ul  ol  h2  h3  h4  table  figure  br
    行内   strong b  em  i  u  s  sup  sub
    结构   li  thead  tbody  tfoot  tr  th  td  figcaption  img

    属性   id（块级，字母开头 ≤64）、table[data-caption]、
           th|td[colspan]、img[src|alt]

**明确不支持**（越界报错，而不是生成半坏的产物）：

    h1                章节标题由 sections[].title 提供，正文请用 h2~h4
    rowspan           一期不做；宁可报错也不生成错位的表格
    ol[start]         一期不支持编号起始值
    style / class     不支持内联样式，字号颜色交给导出模板
    script style iframe svg object embed link meta form input    一律禁止
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from docx.oxml import OxmlElement
from lxml import html as lxml_html

from .settings import settings
from . import bid_template


class HtmlError(ValueError):
    """受限 HTML 的可预期错误（对应 HTTP 422）。"""


INLINE_TAGS = ("strong", "b", "em", "i", "u", "s", "sup", "sub")
HEADING_TAGS = {"h2": 2, "h3": 3, "h4": 4}
BLOCK_TAGS = ("p", "ul", "ol", "h2", "h3", "h4", "table", "figure")
STRUCTURE_TAGS = ("li", "thead", "tbody", "tfoot", "tr", "th", "td", "figcaption", "img", "br")

ALLOWED_TAGS = set(BLOCK_TAGS) | set(HEADING_TAGS) | set(INLINE_TAGS) | set(STRUCTURE_TAGS)
TOP_LEVEL_TAGS = set(BLOCK_TAGS) | {"br"}

# 标签 -> 允许的父标签（未列出表示可作顶层块）
PARENTS = {
    "li": {"ul", "ol"},
    "thead": {"table"}, "tbody": {"table"}, "tfoot": {"table"},
    "tr": {"table", "thead", "tbody", "tfoot"},
    "th": {"tr"}, "td": {"tr"},
    "figcaption": {"figure"},
    "img": {"figure"},
}

ALLOWED_ATTRS = {
    "id": {"p", "ul", "ol", "h2", "h3", "h4", "table", "figure", "li", "tr", "th", "td"},
    "data-caption": {"table"},
    "colspan": {"th", "td"},
    "src": {"img"},
    "alt": {"img"},
}

# 常见越界写法的针对性提示（按 (标签, 属性) 或 ("*", 属性) 命中）
ATTR_HINTS = {
    ("th", "rowspan"): "一期不支持 rowspan；请拆成多行，或改用 colspan",
    ("td", "rowspan"): "一期不支持 rowspan；请拆成多行，或改用 colspan",
    ("ol", "start"): "一期不支持 ol 的编号起始值（start）",
    ("*", "style"): "不支持内联 style；字号/颜色/间距请交给导出模板统一控制",
    ("*", "class"): "不支持 class；样式请交给导出模板统一控制",
}

UNSUPPORTED_TAGS = {
    "h1": "章节标题由 sections[].title 提供，正文里请改用 h2 ~ h4",
}
FORBIDDEN_TAGS = ("script", "style", "iframe", "object", "embed", "link", "meta", "form", "input", "svg")

MAX_HTML_CHARS = 200_000
MAX_ISSUES = 20
FIGURE_SUFFIXES = {".png", ".jpg", ".jpeg"}
_ID_PATTERN = re.compile(r"^[A-Za-z][A-Za-z0-9_-]{0,63}$")

EXAMPLE_HTML = (
    '<p>本节说明<strong>实施要点</strong>与<em>质量目标</em>。</p>'
    '<h2>组织架构</h2>'
    '<ul><li>项目经理 1 人</li><li>实施工程师 3 人</li></ul>'
    '<table data-caption="人员配置"><thead><tr><th>岗位</th><th>人数</th></tr></thead>'
    '<tbody><tr><td>项目经理</td><td>1</td></tr></tbody></table>'
    '<figure><img src="flow.png" alt="评标流程"/><figcaption>评标流程</figcaption></figure>'
)


# ---------- 校验 ----------


def _result(issues: list[dict], stats: dict) -> dict:
    return {"ok": not any(item["level"] == "error" for item in issues),
            "issues": issues, "stats": stats}


def validate_html(html: Any) -> dict:
    """校验受限 HTML。返回 {ok, issues, stats}；issues 带精确定位信息便于模型改对。"""
    text = str(html or "")
    issues: list[dict] = []
    stats: dict = {"chars": len(text), "tags": {}, "images": 0, "tables": 0}

    def add(kind: str, message: str, tag: str = "", attr: str = "", level: str = "error") -> None:
        if len(issues) < MAX_ISSUES:
            issues.append({"level": level, "kind": kind, "message": message, "tag": tag, "attr": attr})

    if not text.strip():
        add("empty", "html 为空")
        return _result(issues, stats)
    if len(text) > MAX_HTML_CHARS:
        add("too_long", f"html 过长（{len(text)} 字符，上限 {MAX_HTML_CHARS}）")
        return _result(issues, stats)

    lowered = text.lower()
    for tag in FORBIDDEN_TAGS:
        if f"<{tag}" in lowered:
            add("forbidden_tag", f"不允许出现 <{tag}>", tag=tag)

    try:
        root = lxml_html.fragment_fromstring(text, create_parent="div")
    except Exception as exc:  # noqa: BLE001
        add("parse", f"HTML 解析失败：{type(exc).__name__}: {str(exc)[:120]}")
        return _result(issues, stats)

    def walk(node, parent_tag):
        for child in node:
            if not isinstance(child.tag, str):       # 注释等非元素节点
                continue
            tag = child.tag
            stats["tags"][tag] = stats["tags"].get(tag, 0) + 1

            if tag in UNSUPPORTED_TAGS:
                add("unsupported_tag", f"<{tag}>：{UNSUPPORTED_TAGS[tag]}", tag=tag)
            elif tag not in ALLOWED_TAGS:
                add("unknown_tag", f"不支持的标签 <{tag}>", tag=tag)
            else:
                allowed_parents = PARENTS.get(tag)
                if parent_tag is None:
                    if tag not in TOP_LEVEL_TAGS:
                        add("misplaced", f"<{tag}> 不能作为顶层块，只能出现在它的容器内部", tag=tag)
                elif allowed_parents is not None and parent_tag not in allowed_parents:
                    add("misplaced",
                        f"<{tag}> 只能出现在 {'/'.join(sorted(allowed_parents))} 内，当前在 <{parent_tag}> 里",
                        tag=tag)

            for attr, value in (child.attrib or {}).items():
                hint = ATTR_HINTS.get((tag, attr)) or ATTR_HINTS.get(("*", attr))
                if hint:
                    add("unsupported_attr", f"<{tag}> 的 {attr}：{hint}", tag=tag, attr=attr)
                    continue
                owners = ALLOWED_ATTRS.get(attr)
                if not owners or tag not in owners:
                    add("unknown_attr", f"<{tag}> 不支持属性 {attr}", tag=tag, attr=attr)
                    continue
                if attr == "id" and not _ID_PATTERN.fullmatch(str(value)):
                    add("bad_id", f"id 不合法（字母开头、≤64、仅字母数字与 _ -）：{value!r}",
                        tag=tag, attr=attr)
                if attr == "colspan":
                    try:
                        if not 1 <= int(value) <= 20:
                            raise ValueError
                    except (TypeError, ValueError):
                        add("bad_colspan", f"colspan 必须是 1~20 的整数：{value!r}", tag=tag, attr=attr)

            if tag == "img":
                stats["images"] += 1
                if not str(child.get("src") or "").strip():
                    add("img_no_src", "<img> 缺少 src（应为 fig-render 产出的图片名）", tag=tag)
            elif tag == "table":
                stats["tables"] += 1

            walk(child, tag)

    walk(root, None)

    for figure in root.iter("figure"):
        if figure.find("img") is None:
            add("figure_no_img", "<figure> 里缺少 <img>", tag="figure")

    return _result(issues, stats)


def html_spec() -> dict:
    """受限 HTML 的规格说明（供模型照着写）。"""
    return {"block_tags": list(BLOCK_TAGS), "heading_tags": list(HEADING_TAGS),
            "inline_tags": list(INLINE_TAGS), "structure_tags": list(STRUCTURE_TAGS),
            "attributes": {attr: sorted(tags) for attr, tags in ALLOWED_ATTRS.items()},
            "not_supported": {
                "h1": UNSUPPORTED_TAGS["h1"],
                "rowspan": "一期不支持；请拆成多行或改用 colspan",
                "ol[start]": "一期不支持编号起始值",
                "style / class": "不支持内联样式与 class；排版交给导出模板",
                "script/style/iframe/svg/object/embed/link/meta/form/input": "一律禁止",
            },
            "example": EXAMPLE_HTML,
            "rules": [
                "章节标题由 sections[].title 提供；正文子标题用 h2 ~ h4。",
                "图片必须包在 figure 里：<figure><img src=\"flow.png\" alt=\"…\"/><figcaption>图题</figcaption></figure>；"
                "src 用 fig-render 产出的图片名。",
                "表格用 table[data-caption] 给表题；表头行用 thead 里的 th。",
                "越界会被**精确报错**（不静默删掉），先调 html_validate 自查再出稿。",
            ]}


# ---------- 渲染 ----------


def resolve_figure_path(raw: Any) -> Path | None:
    """解析图片路径：绝对路径直接用；相对名优先在 <BID_OUTPUT_DIR>/figures/ 下找。"""
    text = str(raw or "").strip()
    if not text:
        return None
    candidate = Path(text)
    candidates = [candidate] if candidate.is_absolute() else [
        Path(settings.BID_OUTPUT_DIR) / "figures" / text,
        Path(settings.BID_OUTPUT_DIR) / text,
    ]
    for path in candidates:
        if path.suffix.lower() in FIGURE_SUFFIXES and path.is_file():
            return path
    return None


def _add_run(paragraph, text: str, flags: dict):
    if not text:
        return None
    run = paragraph.add_run(text)
    font = run.font
    if flags.get("bold"):
        font.bold = True
    if flags.get("italic"):
        font.italic = True
    if flags.get("underline"):
        font.underline = True
    if flags.get("strike"):
        font.strike = True
    if flags.get("sup"):
        font.superscript = True
    if flags.get("sub"):
        font.subscript = True
    return run


def _emit_inline(paragraph, node, flags: dict | None = None) -> None:
    """把节点及其子节点渲染成行内 run（递归处理 strong/em/u/s/sup/sub/br）。"""
    flags = flags or {}
    _add_run(paragraph, node.text, flags)
    for child in node:
        if not isinstance(child.tag, str):
            continue
        tag = child.tag
        if tag == "br":
            paragraph.add_run().add_break()
        else:
            merged = dict(flags)
            merged["bold"] = flags.get("bold") or tag in ("strong", "b")
            merged["italic"] = flags.get("italic") or tag in ("em", "i")
            merged["underline"] = flags.get("underline") or tag == "u"
            merged["strike"] = flags.get("strike") or tag == "s"
            merged["sup"] = flags.get("sup") or tag == "sup"
            merged["sub"] = flags.get("sub") or tag == "sub"
            _emit_inline(paragraph, child, merged)
        _add_run(paragraph, child.tail, flags)


def _colspan(cell) -> int:
    try:
        return max(1, min(20, int(cell.get("colspan") or 1)))
    except (TypeError, ValueError):
        return 1


def _render_table(doc, node, template: dict, state: dict) -> None:
    grid: list[tuple[Any, list]] = []
    for row in node.iter("tr"):
        cells = [(cell, _colspan(cell)) for cell in row
                 if isinstance(cell.tag, str) and cell.tag in ("th", "td")]
        if cells:
            grid.append((row, cells))
    if not grid:
        raise HtmlError("table 里没有可渲染的 td/th")

    columns = max(sum(span for _, span in cells) for _, cells in grid)
    if columns > 20:
        raise HtmlError(f"表格列数过多（{columns} 列）")

    caption = str(node.get("data-caption") or "").strip()
    if caption:
        state["table"] += 1
        bid_template.add_caption(
            doc, f"{template['table']['caption_prefix']} {state['table']} {caption}",
            size_pt=template['table']['font_size_pt'])

    table = doc.add_table(rows=0, cols=columns)
    table.style = template["table"]["style"]
    header_row = None
    for _row_element, cells in grid:
        row = table.add_row()
        index = 0
        for cell_element, span in cells:
            target = row.cells[index]
            _emit_inline(target.paragraphs[0], cell_element)
            end = min(index + span, columns) - 1
            if end > index:
                target.merge(row.cells[end])
            index += span
        if header_row is None and all(cell.tag == "th" for cell, _ in cells):
            header_row = row
            header_row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    bid_template.style_table(table, template, header_row)


def _render_figure(doc, node, template: dict, state: dict) -> None:
    image = node.find("img")
    if image is None:
        raise HtmlError("<figure> 里缺少 <img>")
    raw = str(image.get("src") or "").strip()
    path = resolve_figure_path(raw)
    if path is None:
        state["missing"].append(f"配图未找到：{raw[:80]}")
        return
    caption_element = node.find("figcaption")
    title = "".join(caption_element.itertext()).strip() if caption_element is not None else ""
    if not title:
        title = str(image.get("alt") or "").strip()
    caption = ""
    if title:
        state["figure"] += 1
        caption = f"{template['image']['caption_prefix']} {state['figure']} {title}"
    bid_template.add_figure(doc, path, template, caption=caption)


def _render_block(doc, node, template: dict, state: dict) -> None:
    tag = node.tag
    if tag in HEADING_TAGS:
        _emit_inline(doc.add_paragraph(style=f"Heading {HEADING_TAGS[tag]}"), node)
    elif tag in ("ul", "ol"):
        style = "List Bullet" if tag == "ul" else "List Number"
        for item in node.findall("li"):
            _emit_inline(doc.add_paragraph(style=style), item)
    elif tag == "table":
        _render_table(doc, node, template, state)
    elif tag == "figure":
        _render_figure(doc, node, template, state)
    elif tag == "br":
        doc.add_paragraph()
    else:                                   # p 及其它已放行的块级标签
        _emit_inline(doc.add_paragraph(), node)


def render_html(doc, html: Any, template: dict, *, figure_number: int = 0, table_number: int = 0,
                missing: list[str] | None = None) -> dict:
    """把受限 HTML 渲染进文档；校验不通过直接抛错（不静默降级）。

    返回 {figure_number, table_number, missing, stats}，便于多节之间连续编号。
    """
    report = validate_html(html)
    if not report["ok"]:
        detail = "；".join(item["message"] for item in report["issues"][:5])
        raise HtmlError(f"受限 HTML 校验未通过（{len(report['issues'])} 处）：{detail}")
    root = lxml_html.fragment_fromstring(str(html), create_parent="div")
    state = {"figure": figure_number, "table": table_number,
             "missing": missing if missing is not None else []}
    for node in root:
        if isinstance(node.tag, str):
            _render_block(doc, node, template, state)
    return {"figure_number": state["figure"], "table_number": state["table"],
            "missing": state["missing"], "stats": report["stats"]}


__all__ = ["HtmlError", "INLINE_TAGS", "BLOCK_TAGS", "HEADING_TAGS", "STRUCTURE_TAGS",
           "ALLOWED_TAGS", "ALLOWED_ATTRS", "validate_html", "html_spec",
           "render_html", "resolve_figure_path"]
