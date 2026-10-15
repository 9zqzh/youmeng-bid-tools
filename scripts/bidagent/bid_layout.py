"""版面估算与字数配额：出稿前发现「字数离目标太远 / 末页大半空白」这类明显问题。

**这是确定性估算，不是 Word 实际分页。** 由模板的纸张、页边距、正文字号与行距推出
「每行字数 / 每页行数」，再按章节的字数、图片（读实际像素、按模板画框等比换算高度）、
表格（按行数估高）累加，得到估算页数与**末页填充率**。

为什么不直接渲染测量：精确分页要 Word/LibreOffice 排版引擎，还要多一次往返与依赖。
先做确定性估算，把明显不合理的配额与版面挡在出稿前，性价比更高 —— 但它**不能替代**
在 Word 里核对版式（这一点也在返回里明确写了）。

三块产出：
- **字数配额**：给总字数 → 按 `weight` 分配到各节（带上下限并自动再分配）；给每节目标 → 照用；
  某节显式带 `target_words` → 该节优先（其余节再分剩余）。
- **版面估算**：每节估算页数、末页填充率、图片与表格占位；图片高度用**真实像素**换算。
- **偏差与建议**：字数偏离超过阈值、末页填充率过低 → 给出可执行动作（扩写/缩写/挪图）。
"""
from __future__ import annotations

import math
import re
from typing import Any

from lxml import html as lxml_html

from . import bid_html, bid_template


class LayoutError(ValueError):
    """版面模块的可预期错误（对应 HTTP 422）。"""


# 换算与经验标定（都可按实际观感微调）
CM_PER_PT = 0.03528
LINE_FACTOR = 1.30          # Word 中文单倍行距 ≈ 字号 × 1.3
ROW_PADDING_CM = 0.14       # 表格单元格上下内边距估算
CAPTION_CM = 0.65           # 单条题注高度估算
PARAGRAPH_TAIL_LINES = 0.5  # 每个段落末行的「半行」浪费估算

DEVIATION_THRESHOLD = 0.20  # 字数偏离超过 20% 才提示
GAP_THRESHOLD = 0.35        # 末页填充率低于 35% 视为页尾空洞

ESTIMATE_NOTE = ("版面为**估算值**（按模板度量与真实图片尺寸推算），用于配额分配与出稿前自查，"
                 "不等同于 Word 实际分页；定稿前仍需在 Word 里核对版式。")


# ---------- 模板度量 ----------


def page_metrics(template: dict) -> dict:
    """由模板推出每行字数 / 每页行数 / 每页字数等度量。"""
    page = template["page"]
    body = template["body"]
    width, height = bid_template.PAPER_SIZES[page["paper"]]
    if page["orientation"] == "landscape":
        width, height = height, width
    margins = page["margins_cm"]
    usable_width = width - margins["left"] - margins["right"]
    usable_height = height - margins["top"] - margins["bottom"]

    size_pt = body["size_pt"]
    char_width = size_pt * CM_PER_PT                       # 中文按方块字估算
    line_height = char_width * LINE_FACTOR * (body["line_spacing"] or 1.0)
    chars_per_line = max(1, int(usable_width / char_width))
    lines_per_page = max(1, int(usable_height / line_height))
    return {
        "paper": page["paper"], "orientation": page["orientation"],
        "usable_cm": [round(usable_width, 2), round(usable_height, 2)],
        "body_size_pt": size_pt, "line_spacing": body["line_spacing"] or 1.0,
        "chars_per_line": chars_per_line, "lines_per_page": lines_per_page,
        "chars_per_page": chars_per_line * lines_per_page,
        "line_height_cm": round(line_height, 3),
        "image_box_cm": [template["image"]["max_width_cm"], template["image"]["max_height_cm"]],
        "note": ESTIMATE_NOTE,
    }


# ---------- 章节度量 ----------


def _count_chars(text: Any) -> int:
    return len(re.sub(r"\s+", "", str(text or "")))


def _html_parts(html: Any) -> tuple[int, list[str], list[int], int]:
    """从受限 HTML 里取：字数、图片引用、每张表的行数、块级元素数。"""
    root = lxml_html.fragment_fromstring(str(html), create_parent="div")
    chars = _count_chars(root.text_content())
    figures = [img.get("src") for img in root.iter("img")]
    tables = [len(list(table.iter("tr"))) for table in root.iter("table")]
    blocks = sum(1 for node in root.iter()
                 if isinstance(node.tag, str) and node.tag in {"p", "li", "h2", "h3", "h4"})
    return chars, figures, tables, max(1, blocks)


def _plain_parts(section: dict) -> tuple[int, list[str], list[int], int]:
    lines = [line for line in str(section.get("content") or "").splitlines() if line.strip()]
    chars = _count_chars(section.get("content"))
    figures = [_figure_ref(item) for item in section.get("images") or []]
    tables = []
    table = section.get("table")
    if isinstance(table, dict):
        tables.append(len(table.get("rows") or []) + 1 + (1 if table.get("title") else 0))
    return chars, figures, tables, max(1, len(lines))


def _figure_ref(item: Any) -> str:
    if isinstance(item, str):
        return item.strip()
    if not isinstance(item, dict):
        return ""
    return str(item.get("file") or item.get("name") or item.get("path") or "").strip()


def _figure_height_cm(path, template: dict) -> float:
    """按模板画框把图片等比缩放后的高度（厘米）—— 用真实像素算，不猜。"""
    from PIL import Image

    with Image.open(path) as image:
        width, height = image.size
    config = template["image"]
    scale = min(config["max_width_cm"] / width, config["max_height_cm"] / height)
    return height * scale


def _table_height_cm(rows: int, template: dict, *, caption: bool) -> float:
    font_pt = template["table"]["font_size_pt"] or template["body"]["size_pt"]
    row_cm = font_pt * LINE_FACTOR * CM_PER_PT + ROW_PADDING_CM * 2
    return rows * row_cm + (CAPTION_CM if caption else 0.0)


def measure_section(section: dict, template: dict, metrics: dict) -> dict:
    """估算单节的字数与版面占位。"""
    if section.get("html"):
        chars, figure_refs, table_row_counts, blocks = _html_parts(section["html"])
        table_captions = str(section["html"]).count("data-caption")
    else:
        chars, figure_refs, table_row_counts, blocks = _plain_parts(section)
        table_captions = 1 if (section.get("table") or {}).get("title") else 0

    missing: list[str] = []
    figure_cm = 0.0
    for reference in figure_refs:
        path = bid_html.resolve_figure_path(reference)
        if path is None:
            missing.append(reference[:80])
            continue
        figure_cm += _figure_height_cm(path, template) + CAPTION_CM

    table_cm = 0.0
    for index, rows in enumerate(table_row_counts):
        table_cm += _table_height_cm(rows, template, caption=index < table_captions)

    text_lines = chars / metrics["chars_per_line"] + blocks * PARAGRAPH_TAIL_LINES
    text_cm = text_lines * metrics["line_height_cm"]
    total_cm = text_cm + figure_cm + table_cm
    usable = metrics["usable_cm"][1]

    raw_pages = total_cm / usable if usable else 0.0
    pages = max(1, math.ceil(raw_pages)) if total_cm > 0 else 0
    fill = 0.0
    if raw_pages > 0:
        fill = round(raw_pages % 1, 2) or 1.0
        if raw_pages < 1:
            fill = round(raw_pages, 2)
    return {"chars": chars, "blocks": blocks, "figures": len(figure_refs) - len(missing),
            "tables": len(table_row_counts), "missing_figures": missing,
            "text_cm": round(text_cm, 1), "figure_cm": round(figure_cm, 1),
            "table_cm": round(table_cm, 1), "total_cm": round(total_cm, 1),
            "estimated_pages": pages, "last_page_fill": fill}


# ---------- 字数配额 ----------


def _proportional(total: int, weights: list[float], low: int, high: int) -> list[int]:
    """按权重分配 total，带 [low, high] 限制；越界的固定后把余额再分给其余项。"""
    count = len(weights)
    if count == 0:
        return []
    result = [0.0] * count
    fixed = [False] * count
    remaining = float(total)
    for _ in range(8):
        free = [i for i in range(count) if not fixed[i]]
        if not free:
            break
        weight_sum = sum(weights[i] for i in free) or float(len(free))
        changed = False
        for index in free:
            share = remaining * ((weights[index] if weight_sum else 1.0) / weight_sum)
            if share < low:
                result[index], fixed[index] = low, True
                remaining -= low
                changed = True
            elif share > high:
                result[index], fixed[index] = high, True
                remaining -= high
                changed = True
        if not changed:
            for index in free:
                result[index] = remaining * ((weights[index] if weight_sum else 1.0) / weight_sum)
            break
    return [max(0, int(round(value))) for value in result]


def allocate(sections: list[dict], *, total_words: int | None = None,
             section_words: int | None = None, min_words: int | None = None,
             max_words: int | None = None) -> tuple[list[int], str]:
    """给出每节的目标字数。返回 (targets, mode)。

    优先用每节显式的 `target_words`；剩下的再按 `weight` 分 `total_words`。
    既没总字数也没每节字数时返回空（表示不做配额）。
    """
    count = len(sections)
    if count == 0:
        return [], "none"
    explicit = [int(item["target_words"]) for item in sections
                if isinstance(item.get("target_words"), (int, float)) and item["target_words"] > 0]
    explicit = explicit if len(explicit) == count else []

    if explicit:
        return explicit, "per_section"
    if section_words:
        return [int(section_words)] * count, "per_section"
    if not total_words:
        return [], "none"

    low = int(min_words) if min_words else max(50, int(total_words) // (count * 5))
    high = int(max_words) if max_words else max(low, int(total_words) // max(1, count // 3))
    weights = [max(0.0, float(item.get("weight") or 1)) for item in sections]
    return _proportional(int(total_words), weights, low, high), "total"


# ---------- 汇总 ----------


def plan(sections: Any, *, template: Any = None, total_words: int | None = None,
         section_words: int | None = None, min_words: int | None = None,
         max_words: int | None = None) -> dict:
    """配额 + 版面估算 + 偏差与建议。"""
    if not isinstance(sections, list) or not sections:
        raise LayoutError("sections 必须是非空数组")
    resolved = bid_template.resolve_template(template)
    metrics = page_metrics(resolved)

    normalized = []
    for index, item in enumerate(sections, 1):
        if not isinstance(item, dict):
            continue
        normalized.append({"id": str(item.get("id") or f"sec_{index:04d}"),
                           "title": str(item.get("title") or "")[:120],
                           "weight": item.get("weight"), "target_words": item.get("target_words"),
                           "html": item.get("html"), "content": item.get("content"),
                           "images": item.get("images"), "table": item.get("table")})
    if not normalized:
        raise LayoutError("sections 里没有可用的小节")

    targets, mode = allocate(normalized, total_words=total_words, section_words=section_words,
                             min_words=min_words, max_words=max_words)

    rows: list[dict] = []
    suggestions: list[str] = []
    under = over = gaps = 0
    total_chars = 0
    for index, section in enumerate(normalized):
        measured = measure_section(section, resolved, metrics)
        total_chars += measured["chars"]
        target = targets[index] if targets else None
        deviation = None
        flags: list[str] = []
        if target:
            deviation = round((measured["chars"] - target) / target, 3)
            if deviation <= -DEVIATION_THRESHOLD:
                flags.append("字数不足")
                under += 1
                suggestions.append("《{}》比目标少 {} 字（当前 {}），建议扩写".format(
                    section["title"] or section["id"], target - measured["chars"], measured["chars"]))
            elif deviation >= DEVIATION_THRESHOLD:
                flags.append("字数超出")
                over += 1
                suggestions.append("《{}》比目标多 {} 字（当前 {}），建议精简".format(
                    section["title"] or section["id"], measured["chars"] - target, measured["chars"]))
        if measured["estimated_pages"] >= 1 and measured["last_page_fill"] < GAP_THRESHOLD:
            flags.append("页尾空洞")
            gaps += 1
            suggestions.append("《{}》末页只填了 {:.0%}，建议调整字数或把图/表挪到相邻小节".format(
                section["title"] or section["id"], measured["last_page_fill"]))
        if measured["missing_figures"]:
            flags.append("配图缺失")
            suggestions.append("《{}》引用的配图不存在：{}".format(
                section["title"] or section["id"], "、".join(measured["missing_figures"])))

        rows.append({"id": section["id"], "title": section["title"],
                     "chars": measured["chars"], "target_words": target, "deviation": deviation,
                     "weight": section.get("weight"),
                     "figures": measured["figures"], "tables": measured["tables"],
                     "missing_figures": measured["missing_figures"],
                     "estimated_pages": measured["estimated_pages"],
                     "last_page_fill": measured["last_page_fill"],
                     "height_cm": measured["total_cm"], "flags": flags})

    pages_total = sum(row["estimated_pages"] for row in rows)
    target_total = sum(row["target_words"] for row in rows if row["target_words"])
    return {
        "metrics": metrics,
        "budget": {"mode": mode, "total_words": total_words, "target_total": target_total or None,
                   "sections": [{"id": row["id"], "title": row["title"], "weight": row["weight"],
                                 "target_words": row["target_words"]} for row in rows]},
        "sections": rows,
        "summary": {"sections": len(rows), "total_chars": total_chars,
                    "target_total": target_total or None,
                    "pages_estimated": pages_total, "under": under, "over": over,
                    "page_gaps": gaps,
                    "ok": not (under or over or gaps or any(row["missing_figures"] for row in rows))},
        "suggestions": suggestions,
        "note": ESTIMATE_NOTE,
    }


__all__ = ["LayoutError", "ESTIMATE_NOTE", "DEVIATION_THRESHOLD", "GAP_THRESHOLD",
           "page_metrics", "measure_section", "allocate", "plan"]
