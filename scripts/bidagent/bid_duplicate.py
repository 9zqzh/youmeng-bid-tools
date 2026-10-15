"""标书查重（确定性引擎）：把「可算的部分」算准，把「要判断的部分」留给模型。

分工的原因：你们既有链路是「后端提供候选 → 模型只解释候选」，而那条 LLM 提示词明确写着
「仅解释给出的候选…没有算法数值就明确无法量化」。也就是说**算法本该由后端算，之前是缺的**。
本模块补的就是这一层。

四层确定性比对 + 一条贯穿性原则：

1. **招标白名单剔除（最重要）** —— 招标原文会被各家投标文件大段照抄（须知、格式、评标办法）。
   若不做剔除，这些合法引用会**大面积误报**成"雷同"。所以先把招标原文建成索引，
   投标文件里命中招标来源的句子单独归类为「招标原文引用」并**不计入雷同统计**。
2. **正文句子级比对** —— 归一化后精确等值 + 二元组 Dice/containment 近重复（分层阈值）。
3. **目录比对** —— 标题用字符 bigram Dice，顺序用 LCS；文件对相似度 = 路径 0.45 + 标题 0.35 + 顺序 0.2。
4. **元数据取证** —— docx 的 rsid 会话指纹与属性、PDF 的 trailer ID / 增量更新次数 / 签名数、
   整文件 sha256：查的是「同源 / 同一份改的」痕迹，而不是文字雷同。
5. **图片哈希** —— docx 内嵌媒体与独立图片文件的字节 sha256，跨文件相同即复用。

**结果失效判定**：返回 `signature`（输入内容与选项的指纹）。输入一变，旧结论就不该再被采信。

确定性输出，不调用任何模型；`candidates` 直接就是既有 Dify 节点要的
`{current_source_id, reference_source_id, similarity, overlap_excerpt}` 形态。
**法定固定格式 vs 实质重复的定性判断留给模型** —— 引擎只给可复核的数值与原文摘录。
"""
from __future__ import annotations

import hashlib
import json
import re
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from .bidding_documents import DocumentError, parse_document


class DuplicateError(ValueError):
    """查重模块的可预期错误（对应 HTTP 422）。"""


# ---------- 可调阈值（对齐 OpenBidKit 的实测取值） ----------

MIN_SENTENCE_CHARS = 20          # 短于这个长度不参与比对（噪音太多）
SHORT_SENTENCE_CHARS = 30        # 短句 / 长句分界
SHORT_NEAR = (0.95, 0.88, 0.55)  # 短句近重复：(containment, dice, length_ratio)
LONG_NEAR = (0.90, 0.82, 0.50)   # 长句近重复
OUTLINE_WEIGHTS = (0.45, 0.35, 0.20)   # 目录文件对：(路径重叠, 标题相似, 顺序相似)
RISK_BANDS = ((0.75, "high"), (0.55, "medium"), (0.35, "low"))
MAX_DOCUMENTS = 12
MAX_CANDIDATES = 200
MAX_IMAGE_BYTES = 20 * 1024 * 1024

_PUNCTUATION = re.compile(r"[\s\u3000!-/:-@\[-`{-~，。、；：？！“”‘’（）《》〈〉【】…—·「」『』]+")
_DIGITS = re.compile(r"\d+")
_SENTENCE_SPLIT = re.compile(r"[。！？；\n]+")
_HEADING = re.compile(
    r"^\s*(?:第\s*[一二三四五六七八九十百零〇\d]+\s*[章节篇部分条]|"
    r"[一二三四五六七八九十]+\s*[、.．]|\d+(?:\.\d+)*\s*[、.．\s])"
)


def _hash_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _hash_text(text: str) -> str:
    return _hash_bytes(str(text).encode("utf-8"))


def normalize_sentence(text: Any) -> str:
    """归一化：全角转半角、去空白与标点、小写。保留数字（数字参与比对有意义）。"""
    converted = unicodedata.normalize("NFKC", str(text or "")).lower()
    return _PUNCTUATION.sub("", converted)


def skeleton_sentence(text: Any) -> str:
    """骨架：在归一化基础上抹掉数字 —— 抓「把工期 60 天改成 90 天」这类模板套用。"""
    return _DIGITS.sub("#", normalize_sentence(text))


def split_sentences(text: Any, *, min_chars: int = MIN_SENTENCE_CHARS) -> list[str]:
    """按句末标点切句，过滤过短与纯符号句。"""
    pieces = []
    for raw in _SENTENCE_SPLIT.split(str(text or "")):
        sentence = raw.strip()
        if len(normalize_sentence(sentence)) >= min_chars:
            pieces.append(sentence)
    return pieces


def _bigrams(text: str) -> set[str]:
    return {text[i:i + 2] for i in range(len(text) - 1)} or ({text} if text else set())


def _dice(left: set[str], right: set[str]) -> float:
    if not left or not right:
        return 0.0
    return 2 * len(left & right) / (len(left) + len(right))


def _containment(left: set[str], right: set[str]) -> float:
    small, large = (left, right) if len(left) <= len(right) else (right, left)
    return len(small & large) / len(small) if small else 0.0


def _lcs_length(left: list, right: list) -> int:
    if not left or not right:
        return 0
    previous = [0] * (len(right) + 1)
    for item in left:
        current = [0]
        for index, other in enumerate(right, 1):
            if item == other:
                current.append(previous[index - 1] + 1)
            else:
                current.append(max(previous[index], current[-1]))
        previous = current
    return previous[-1]


# ---------- 招标白名单 ----------


def build_tender_index(tender_text: Any) -> dict:
    """把招标原文建成多级索引：精确 / 骨架 / 前缀桶（用于近重复）。"""
    sentences = split_sentences(tender_text)
    exact = {normalize_sentence(item): item for item in sentences}
    skeleton = {skeleton_sentence(item) for item in sentences}
    prefixes = defaultdict(list)
    for sentence in sentences:
        normalized = normalize_sentence(sentence)
        prefixes[normalized[:10]].append(normalized)
    return {"exact": exact, "skeleton": skeleton, "prefixes": dict(prefixes),
            "sentences": len(sentences), "chars": len(str(tender_text or ""))}


def classify_against_tender(normalized: str, index: dict) -> str:
    """判断一句话是否来自招标原文：exact / skeleton / near / none。"""
    if not index or not index.get("exact"):
        return "none"
    if normalized in index["exact"]:
        return "exact"
    if skeleton_sentence(normalized) in index["skeleton"]:
        return "skeleton"
    candidates = index["prefixes"].get(normalized[:10]) or []
    target = _bigrams(normalized)
    for other in candidates:
        containment, dice = _containment(target, _bigrams(other)), _dice(target, _bigrams(other))
        if containment >= SHORT_NEAR[0] and dice >= SHORT_NEAR[1]:
            return "near"
    return "none"


# ---------- 正文比对 ----------


def _prepare_document(document: dict) -> dict:
    identifier = str(document.get("id") or document.get("name") or "")
    name = str(document.get("name") or identifier)[:120]
    text = document.get("text")
    source_path = document.get("path")
    notes: list[str] = []
    if text is None and source_path:
        path = Path(str(source_path))
        if not path.is_file():
            raise DuplicateError(f"文档不存在：{path}")
        try:
            result = parse_document(path)
        except DocumentError as exc:
            raise DuplicateError(f"解析文档失败（{path.name}）：{exc}") from exc
        text = "\n".join(block["text"] for block in result.get("blocks", []))
        notes.extend(result.get("warnings", []))
    if not text:
        raise DuplicateError(f"文档 {name or identifier} 既没有 text 也没有可解析的 path")
    return {"id": identifier, "name": name, "text": str(text), "path": str(source_path) if source_path else None,
            "notes": notes}


def read_document_text(path: Any) -> str:
    """读取服务器上的文档并解析为纯文本（PDF/DOCX，含 OCR 回退）。

    供「不想把整篇正文塞进工具参数」的调用方使用（大文档走路径更稳）。
    """
    return _prepare_document({"path": str(path)})["text"]


def detect_content_duplicates(documents: list[dict], tender_index: dict) -> dict:
    """跨文件句子级雷同检测（已剔除招标原文引用）。"""
    entries: list[dict] = []
    tender_hits: list[dict] = []
    by_key: dict[str, list[dict]] = defaultdict(list)

    for document in documents:
        for sentence in split_sentences(document["text"]):
            normalized = normalize_sentence(sentence)
            verdict = classify_against_tender(normalized, tender_index)
            if verdict != "none":
                tender_hits.append({"document_id": document["id"], "verdict": verdict,
                                    "excerpt": sentence[:120]})
                continue
            entry = {"document_id": document["id"], "normalized": normalized,
                     "bigrams": _bigrams(normalized), "excerpt": sentence[:120], "exact_key": normalized}
            entries.append(entry)
            by_key[normalized].append(entry)

    groups: list[dict] = []
    consumed: set[int] = set()

    # 1) 精确等值（同一归一化句子出现在多于一份文件里）
    for normalized, hits in by_key.items():
        files = sorted({hit["document_id"] for hit in hits})
        if len(files) > 1:
            for hit in hits:
                consumed.add(id(hit))
            groups.append({"kind": "exact", "similarity": 1.0, "files": files,
                           "normalized_length": len(normalized),
                           "excerpt": hits[0]["excerpt"]})

    # 2) 近重复：倒排 bigram 找候选，再按分层阈值确认
    inverted: dict[str, list[int]] = defaultdict(list)
    for index, entry in enumerate(entries):
        if id(entry) in consumed:
            continue
        for gram in entry["bigrams"]:
            inverted[gram].append(index)

    for index, entry in enumerate(entries):
        if id(entry) in consumed:
            continue
        counter: Counter[int] = Counter()
        for gram in entry["bigrams"]:
            for other in inverted.get(gram, ()):
                if other != index:
                    counter[other] += 1
        best = None
        for other, _shared in counter.most_common(20):
            candidate = entries[other]
            if id(candidate) in consumed or candidate["document_id"] == entry["document_id"]:
                continue
            ratio = min(len(entry["normalized"]), len(candidate["normalized"])) / \
                max(len(entry["normalized"]), len(candidate["normalized"]))
            containment = _containment(entry["bigrams"], candidate["bigrams"])
            dice = _dice(entry["bigrams"], candidate["bigrams"])
            longer = max(len(entry["normalized"]), len(candidate["normalized"]))
            threshold = SHORT_NEAR if longer < SHORT_SENTENCE_CHARS else LONG_NEAR
            if containment >= threshold[0] and dice >= threshold[1] and ratio >= threshold[2]:
                score = round((containment + dice) / 2, 3)
                if best is None or score > best["similarity"]:
                    best = {"kind": "near", "similarity": score,
                            "files": sorted({entry["document_id"], candidate["document_id"]}),
                            "normalized_length": longer, "excerpt": entry["excerpt"],
                            "counterpart_excerpt": candidate["excerpt"]}
        if best:
            consumed.add(id(entry))
            groups.append(best)

    groups.sort(key=lambda item: (-item["similarity"], -item["normalized_length"]))
    duplicated_sentences = sum(1 for entry in entries if id(entry) in consumed)
    return {"groups": groups, "tender_hits": tender_hits,
            "summary": {"sentences_compared": len(entries), "tender_excluded": len(tender_hits),
                        "duplicate_groups": len(groups), "duplicate_sentences": duplicated_sentences,
                        "duplicate_ratio": round(duplicated_sentences / len(entries), 3) if entries else 0.0}}


# ---------- 目录比对 ----------


def extract_outline(text: Any) -> list[str]:
    """从正文里抽章节标题（启发式）：识别「第X章 / 一、 / 1.1」这类开头。"""
    outline: list[str] = []
    for line in str(text or "").splitlines():
        stripped = line.strip()
        if 2 <= len(stripped) <= 60 and _HEADING.match(stripped):
            outline.append(stripped)
    return outline


def outline_similarity(left: list[str], right: list[str]) -> dict:
    """目录相似度：路径重叠 0.45 + 标题相似 0.35 + 顺序相似 0.2（对齐既有取值）。"""
    left_norm = [normalize_sentence(item) for item in left]
    right_norm = [normalize_sentence(item) for item in right]
    left_set, right_set = set(left_norm), set(right_norm)
    path_overlap = len(left_set & right_set) / min(len(left_set), len(right_set)) if left_set and right_set else 0.0

    if not left_norm or not right_norm:
        title_score = 0.0
    else:
        hits = 0
        for item in left_norm:
            target = _bigrams(item)
            if any(_dice(target, _bigrams(other)) >= 0.86 for other in right_norm):
                hits += 1
        title_score = hits / len(left_norm)

    order_score = _lcs_length(left_norm, right_norm) / min(len(left_norm), len(right_norm)) \
        if left_norm and right_norm else 0.0
    score = (path_overlap * OUTLINE_WEIGHTS[0] + title_score * OUTLINE_WEIGHTS[1]
             + order_score * OUTLINE_WEIGHTS[2])
    return {"path_overlap": round(path_overlap, 3), "title_similarity": round(title_score, 3),
            "order_similarity": round(order_score, 3), "score": round(score, 3),
            "risk": _risk_of(score)}


def _risk_of(score: float) -> str:
    for threshold, label in RISK_BANDS:
        if score >= threshold:
            return label
    return "none"


# ---------- 元数据取证 ----------


def _docx_metadata(path: Path) -> dict:
    from zipfile import ZipFile

    info: dict = {"kind": "docx"}
    with ZipFile(path) as archive:
        names = set(archive.namelist())
        if "word/settings.xml" in names:
            settings = archive.read("word/settings.xml").decode("utf-8", "ignore")
            # 两种常见形式：段落上的 w:rsidR/w:rsidP 属性，以及设置里的 <w:rsid w:val="…"/>
            found = set(re.findall(r'w:rsid\w*="([0-9A-Fa-f]{8})"', settings))
            found |= set(re.findall(r'<w:rsid\s+w:val="([0-9A-Fa-f]{8})"', settings))
            info["rsids"] = sorted(found)
        for part, keys in (("docProps/core.xml", ("creator", "lastModifiedBy", "revision", "created", "modified")),
                           ("docProps/app.xml", ("Application", "Company", "Template"))):
            if part not in names:
                continue
            xml = archive.read(part).decode("utf-8", "ignore")
            for key in keys:
                found = re.search(rf"<[^>]*{key}[^>]*>([^<]*)<", xml, re.I)
                if found and found.group(1).strip():
                    info[key.lower()] = found.group(1).strip()[:80]
        media = [name for name in names if name.startswith("word/media/")]
        info["media_count"] = len(media)
    return info


def _pdf_metadata(path: Path) -> dict:
    data = path.read_bytes()
    info: dict = {"kind": "pdf", "incremental_updates": data.count(b"%%EOF"),
                  "has_signature": b"/Type /Sig" in data or b"/Type/Sig" in data}
    trailer = re.search(rb"/ID\s*\[\s*<([0-9A-Fa-f]+)>", data)
    if trailer:
        info["trailer_id"] = trailer.group(1).decode("ascii", "ignore")
    for key in (b"/Producer", b"/Creator"):
        found = re.search(key + rb"\s*\(([^)]{0,80})\)", data)
        if found:
            info[key.decode().lstrip("/").lower()] = found.group(1).decode("utf-8", "ignore").strip()
    return info


def _ole_metadata(path: Path) -> dict:
    data = path.read_bytes()
    return {"kind": "ole", "magic": data[:8].hex(), "streams_guess": data.count(b"\x00\x00\x00\x00")}


def collect_metadata(path_text: str) -> dict:
    path = Path(path_text)
    suffix = path.suffix.lower()
    try:
        if suffix == ".docx":
            info = _docx_metadata(path)
        elif suffix == ".pdf":
            info = _pdf_metadata(path)
        elif suffix == ".doc":
            info = _ole_metadata(path)
        else:
            info = {"kind": suffix.lstrip(".") or "unknown"}
    except Exception as exc:      # noqa: BLE001 - 取证是「尽力而为」：任何解析异常都不该让整次查重失败
        return {"kind": "error", "error": f"{type(exc).__name__}: {str(exc)[:80]}"}
    info["sha256"] = _hash_bytes(path.read_bytes())
    info["size"] = path.stat().st_size
    return info


def detect_same_origin(metadatas: list[dict]) -> dict:
    """同源嫌疑：整文件哈希相同、rsid 会话指纹高度重叠、PDF trailer ID 相同。"""
    findings: list[dict] = []
    by_hash: dict[str, list[str]] = defaultdict(list)
    for item in metadatas:
        if item.get("sha256"):
            by_hash[item["sha256"]].append(item["document_id"])
    for digest, files in by_hash.items():
        if len(files) > 1:
            findings.append({"kind": "identical_file", "files": sorted(files),
                             "detail": f"文件字节完全相同（sha256 {digest[:12]}…）"})

    for index, left in enumerate(metadatas):
        for right in metadatas[index + 1:]:
            left_ids, right_ids = set(left.get("rsids") or []), set(right.get("rsids") or [])
            if left_ids and right_ids:
                shared = left_ids & right_ids
                ratio = len(shared) / min(len(left_ids), len(right_ids))
                if ratio >= 0.5 and len(shared) >= 5:
                    findings.append({"kind": "shared_rsid", "files": sorted({left["document_id"], right["document_id"]}),
                                     "ratio": round(ratio, 3), "shared": len(shared),
                                     "detail": f"编辑会话指纹重叠 {ratio:.0%}（{len(shared)} 个 rsid）——同源或同模板改写痕迹"})
            if left.get("trailer_id") and left.get("trailer_id") == right.get("trailer_id"):
                findings.append({"kind": "same_pdf_trailer", "files": sorted({left["document_id"], right["document_id"]}),
                                 "detail": "PDF trailer ID 相同，通常由同一份文件另存而来"})
    return {"documents": metadatas, "findings": findings}


# ---------- 图片复用 ----------


def collect_images(document: dict) -> list[dict]:
    path_text = document.get("path")
    if not path_text:
        return []
    path = Path(path_text)
    if not path.is_file():
        return []
    results: list[dict] = []
    try:
        if path.suffix.lower() == ".docx":
            from zipfile import ZipFile
            with ZipFile(path) as archive:
                for name in archive.namelist():
                    if not name.startswith("word/media/"):
                        continue
                    data = archive.read(name)
                    if 0 < len(data) <= MAX_IMAGE_BYTES:
                        results.append({"document_id": document["id"], "name": name, "sha256": _hash_bytes(data),
                                        "size": len(data)})
        elif path.suffix.lower() in {".png", ".jpg", ".jpeg"} and path.stat().st_size <= MAX_IMAGE_BYTES:
            data = path.read_bytes()
            results.append({"document_id": document["id"], "name": path.name, "sha256": _hash_bytes(data),
                            "size": len(data)})
    except Exception:      # noqa: BLE001 - 同 collect_metadata：图片提取失败不影响其它层
        return []
    return results


def detect_duplicate_images(images: list[dict]) -> dict:
    by_hash: dict[str, list[dict]] = defaultdict(list)
    for image in images:
        by_hash[image["sha256"]].append(image)
    reusable = []
    for digest, items in by_hash.items():
        files = sorted({item["document_id"] for item in items})
        if len(files) > 1:
            reusable.append({"sha256": digest, "files": files,
                             "count": len(items), "size": items[0]["size"],
                             "detail": f"同一张图（sha256 {digest[:12]}…）出现在 {len(files)} 份文件中"})
    reusable.sort(key=lambda item: -item["count"])
    return {"images": len(images), "reused": reusable}


# ---------- 对外入口 ----------


def _normalize_documents(raw: Any) -> list[dict]:
    if not isinstance(raw, list) or not raw:
        raise DuplicateError("documents 必须是非空数组")
    if len(raw) > MAX_DOCUMENTS:
        raise DuplicateError(f"单次最多比对 {MAX_DOCUMENTS} 份文件（当前 {len(raw)}），请分批")
    prepared = []
    seen: set[str] = set()
    for index, item in enumerate(raw, 1):
        if isinstance(item, str):
            item = {"text": item}
        if not isinstance(item, dict):
            continue
        document = _prepare_document(item)
        if not document["id"]:
            document["id"] = f"doc_{index:04d}"
        if document["id"] in seen:
            document["id"] = f"doc_{index:04d}"
        seen.add(document["id"])
        prepared.append(document)
    if len(prepared) < 2:
        raise DuplicateError("至少需要两份可比对的文件")
    return prepared


def analyze(payload: dict) -> dict:
    """执行一次确定性查重。`tender_text`/`tender_path` 用于白名单剔除。"""
    if not isinstance(payload, dict):
        raise DuplicateError("payload 必须是对象")
    documents = _normalize_documents(payload.get("documents"))

    tender_text = payload.get("tender_text")
    tender_path = payload.get("tender_path")
    warnings: list[str] = []
    if tender_text is None and tender_path:
        tender_file = Path(str(tender_path))
        if not tender_file.is_file():
            raise DuplicateError(f"招标文件不存在：{tender_file}")
        try:
            parsed = parse_document(tender_file)
        except DocumentError as exc:
            raise DuplicateError(f"解析招标文件失败（{tender_file.name}）：{exc}") from exc
        tender_text = "\n".join(block["text"] for block in parsed.get("blocks", []))
        warnings.extend(parsed.get("warnings", []))
    tender_index = build_tender_index(tender_text or "")
    if not tender_index["sentences"]:
        warnings.append("未提供招标原文：无法做**招标白名单剔除**，招标须知/格式条款会被误判为雷同，"
                        "结论仅供内部参考。强烈建议传入 tender_text 或 tender_path。")

    content = detect_content_duplicates(documents, tender_index)

    outlines = {document["id"]: (document.get("outline") or extract_outline(document["text"]))
                for document in documents}
    pairs = []
    for index, left in enumerate(documents):
        for right in documents[index + 1:]:
            similarity = outline_similarity(outlines[left["id"]], outlines[right["id"]])
            if similarity["score"] > 0:
                pairs.append({"left": left["id"], "right": right["id"],
                              "left_outline_size": len(outlines[left["id"]]),
                              "right_outline_size": len(outlines[right["id"]]), **similarity})
    pairs.sort(key=lambda item: -item["score"])

    with_path = [document for document in documents if document.get("path")]
    metadata = {"documents": [], "findings": [],
                "note": "未提供文件路径，跳过元数据取证与图片哈希（需要真实文件字节）"}
    images = {"images": 0, "reused": []}
    if with_path:
        metadatas = []
        for document in with_path:
            info = collect_metadata(document["path"])
            info["document_id"] = document["id"]
            info["name"] = document["name"]
            metadatas.append(info)
        metadata = detect_same_origin(metadatas)
        images = detect_duplicate_images([image for document in with_path for image in collect_images(document)])
        if len(with_path) < len(documents):
            metadata["note"] = f"仅 {len(with_path)}/{len(documents)} 份提供了文件路径，元数据取证只覆盖这些"

    candidates = _build_candidates(content, pairs)
    signature = _signature(payload, documents, tender_text)

    summary = {
        "documents": len(documents),
        "with_file": len(with_path),
        "tender_sentences": tender_index["sentences"],
        "sentences_compared": content["summary"]["sentences_compared"],
        "tender_excluded": content["summary"]["tender_excluded"],
        "duplicate_groups": content["summary"]["duplicate_groups"],
        "duplicate_ratio": content["summary"]["duplicate_ratio"],
        "outline_pairs_high": sum(1 for pair in pairs if pair["risk"] == "high"),
        "same_origin_findings": len(metadata.get("findings") or []),
        "reused_images": len(images.get("reused") or []),
        "candidates": len(candidates),
    }
    summary["ok"] = not (content["groups"] or metadata.get("findings") or images.get("reused")
                         or any(pair["risk"] == "high" for pair in pairs))
    if not tender_index["sentences"]:
        summary["ok"] = False
    return {"signature": signature, "summary": summary, "content": content, "outline": {"pairs": pairs},
            "metadata": metadata, "images": images, "candidates": candidates,
            "tender_whitelist": {"sentences": tender_index["sentences"]},
            "warnings": warnings,
            "note": ("本结果是**确定性算法**产出（可复核、可复现）：数值与原文摘录可以直接引用；"
                     "但「法定固定格式 vs 实质重复」这类定性判断需要模型结合上下文解释，"
                     "相似度不等于法律侵权或围串标证据。")}


def _build_candidates(content: dict, pairs: list[dict]) -> list[dict]:
    """整理成既有链路要的候选形态：{current_source_id, reference_source_id, similarity, overlap_excerpt}。"""
    candidates: list[dict] = []
    for group in content["groups"]:
        files = group["files"]
        for index in range(len(files) - 1):
            candidates.append({
                "current_source_id": files[index], "reference_source_id": files[index + 1],
                "similarity": group["similarity"], "kind": group["kind"],
                "overlap_excerpt": group["excerpt"], "category": "实质重复嫌疑"})
    for pair in pairs:
        if pair["risk"] in ("high", "medium"):
            candidates.append({
                "current_source_id": pair["left"], "reference_source_id": pair["right"],
                "similarity": pair["score"], "kind": "outline",
                "overlap_excerpt": f"目录相似度 {pair['score']:.0%}"
                                   f"（路径 {pair['path_overlap']:.0%} / 标题 {pair['title_similarity']:.0%}"
                                   f" / 顺序 {pair['order_similarity']:.0%}）",
                "category": "目录结构雷同"})
    candidates.sort(key=lambda item: -item["similarity"])
    return candidates[:MAX_CANDIDATES]


def _signature(payload: dict, documents: list[dict], tender_text: Any) -> str:
    """输入指纹：内容或选项一变，旧结论即失效。"""
    material = {
        "tender": _hash_text(tender_text or ""),
        "documents": sorted([{"id": document["id"], "name": document["name"],
                              "hash": _hash_text(document["text"])} for document in documents],
                            key=lambda item: item["id"]),
        "options": {"min_sentence_chars": MIN_SENTENCE_CHARS, "short": SHORT_NEAR, "long": LONG_NEAR,
                    "outline_weights": OUTLINE_WEIGHTS},
    }
    return _hash_bytes(json.dumps(material, ensure_ascii=False, sort_keys=True).encode("utf-8"))[:32]


__all__ = ["DuplicateError", "MIN_SENTENCE_CHARS", "SHORT_NEAR", "LONG_NEAR", "RISK_BANDS",
           "normalize_sentence", "skeleton_sentence", "split_sentences", "build_tender_index",
           "classify_against_tender", "detect_content_duplicates", "extract_outline",
           "outline_similarity", "collect_metadata", "detect_same_origin", "collect_images",
           "detect_duplicate_images", "analyze", "read_document_text"]
