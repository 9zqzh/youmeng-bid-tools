"""Local document workers. Async callers must use ``asyncio.to_thread``.

No remote services, shell commands, generated company assertions or signatures.
"""
from __future__ import annotations

import io
import json
import os
from pathlib import Path
import subprocess
import tempfile
from zipfile import ZipFile

import httpx
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.table import Table
from docx.text.paragraph import Paragraph
from PIL import Image, ImageOps
from pypdf import PdfReader
import pypdfium2 as pdfium

from . import bid_html, bid_template


class DocumentError(ValueError):
    """A document cannot be safely read or generated with available capabilities."""


VOLUMES = {'business_technical': '商务技术分册', 'price': '报价分册'}


def _execute(command: str | list[str], arguments: list[str], capability: str) -> None:
    # command 可以是单个可执行文件路径（str），也可以是 (解释器 + 固定参数) 列表。
    parts = [command] if isinstance(command, str) else list(command)
    binary = Path(parts[0]).expanduser() if parts and parts[0] else Path('')
    if not parts or not parts[0] or not binary.is_file():
        raise DocumentError(f'{capability} executable is not configured or does not exist')
    try:
        subprocess.run([str(binary.resolve()), *parts[1:], *arguments], check=True, timeout=120,
                       shell=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                       creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    except (OSError, subprocess.SubprocessError) as exc:
        raise DocumentError(f'{capability} failed or timed out') from exc


def _remote_convert(path: Path, service_url: str, token: str = '') -> dict:
    """由 doc-service 将 .doc 转换为 .docx，再在本地解析段落/表格。"""
    headers = {'Authorization': f'Bearer {token}'} if token else {}
    url = service_url.rstrip('/') + '/convert'
    try:
        with httpx.Client(timeout=180, trust_env=False) as client:
            with path.open('rb') as handle:
                response = client.post(url, headers=headers,
                                       files={'file': (path.name, handle, 'application/msword')})
    except httpx.HTTPError as exc:
        raise DocumentError(f'doc-service 转换调用失败: {type(exc).__name__}') from exc
    if response.status_code != 200:
        raise DocumentError(f'doc-service 转换返回 HTTP {response.status_code}')
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temporary:
        temporary.write(response.content)
        converted = Path(temporary.name)
    try:
        return _parse_docx(converted)
    finally:
        converted.unlink(missing_ok=True)


def _remote_ocr(path: Path, service_url: str, token: str = '') -> dict[int, str]:
    """由 doc-service 对 PDF 逐页 OCR，返回 {页码: 文本}。"""
    headers = {'Authorization': f'Bearer {token}'} if token else {}
    url = service_url.rstrip('/') + '/ocr'
    try:
        with httpx.Client(timeout=300, trust_env=False) as client:
            with path.open('rb') as handle:
                response = client.post(url, headers=headers,
                                       files={'file': (path.name, handle, 'application/pdf')})
    except httpx.HTTPError as exc:
        raise DocumentError(f'doc-service OCR 调用失败: {type(exc).__name__}') from exc
    if response.status_code != 200:
        raise DocumentError(f'doc-service OCR 返回 HTTP {response.status_code}')
    try:
        payload = response.json()
    except ValueError as exc:
        raise DocumentError('doc-service OCR 返回非 JSON') from exc
    pages = payload.get('pages') if isinstance(payload, dict) else None
    if not isinstance(pages, list):
        raise DocumentError('doc-service OCR 返回结构异常')
    mapped: dict[int, str] = {}
    for item in pages:
        if not isinstance(item, dict):
            continue
        number, text = item.get('page'), item.get('text')
        if type(number) is int and isinstance(text, str):
            mapped[number] = text.strip()
    return mapped


def _ocr_pages(path: Path, command: str | list[str], page_count: int) -> dict[int, str]:
    with tempfile.TemporaryDirectory(prefix='bid-ocr-') as temporary:
        result = Path(temporary) / 'pages.json'
        _execute(command, [str(path.resolve()), str(result)], 'OCR')
        if not result.is_file() or result.stat().st_size > 32 * 1024 * 1024:
            raise DocumentError('OCR did not produce a bounded JSON result')
        try:
            payload = json.loads(result.read_text(encoding='utf-8-sig'))
            pages = payload['pages']
            if not isinstance(pages, list):
                raise ValueError('pages must be a list')
            mapped = {}
            for item in pages:
                number, text = item['page'], item['text']
                if (type(number) is not int or not 1 <= number <= page_count
                        or number in mapped or not isinstance(text, str)):
                    raise ValueError('invalid or duplicate page')
                mapped[number] = text.strip()
            return mapped
        except (OSError, ValueError, TypeError, KeyError) as exc:
            raise DocumentError('OCR returned invalid page JSON') from exc


def _parse_pdf(path: Path, ocr_command: str | list[str], ocr_remote: str = '',
               remote_token: str = '') -> dict:
    reader = PdfReader(str(path))
    if reader.is_encrypted:
        raise DocumentError('Encrypted PDF is unsupported; provide a decrypted local copy')
    texts = [(page.extract_text() or '').strip() for page in reader.pages]
    if not texts:
        raise DocumentError('PDF contains no pages or text')
    missing = [index + 1 for index, text in enumerate(texts) if not text]
    hybrid_pages = [index + 1 for index, page in enumerate(reader.pages)
                    if texts[index] and len(page.images)]
    warnings = []
    if hybrid_pages:
        warnings.append(f'PDF 第 {hybrid_pages} 页同时含文字和图片；图片文字未覆盖，须人工核对。')
    if missing:
        if ocr_remote:
            recognized = _remote_ocr(path, ocr_remote, remote_token)
            warnings.append('OCR 由 doc-service 完成，文字需人工核对原页，尤其是金额、日期及否决条款。')
        elif ocr_command:
            recognized = _ocr_pages(path, ocr_command, len(texts))
            warnings.append('OCR 文字需人工核对原页，尤其是金额、日期及否决条款。')
        else:
            raise DocumentError(f'OCR required for PDF pages without text: {missing}')
        for number in missing:
            if not recognized.get(number):
                raise DocumentError(f'OCR returned no text for PDF page {number}')
            texts[number - 1] = recognized[number]
    return {'blocks': [{'locator': f'page:{i}', 'text': text}
                       for i, text in enumerate(texts, 1)],
            'warnings': warnings, 'complete': not hybrid_pages}


def _docx_blocks(container, prefix=''):
    paragraphs = tables = 0
    element = container._element.body if hasattr(container._element, 'body') else container._tc
    for child in element.iterchildren():
        if child.tag == qn('w:p'):
            paragraphs += 1
            text = Paragraph(child, container).text.strip()
            if text:
                yield {'locator': f'{prefix}paragraph:{paragraphs}', 'text': text}
        elif child.tag == qn('w:tbl'):
            tables += 1
            table = Table(child, container)
            for index, row in enumerate(table.rows, 1):
                text = '\t'.join(cell.text.strip() for cell in row.cells)
                if text.strip():
                    yield {'locator': f'{prefix}table:{tables}/row:{index}', 'text': text}
                seen = set()
                for column, cell in enumerate(row.cells, 1):
                    if cell._tc in seen:
                        continue
                    seen.add(cell._tc)
                    if cell.tables:
                        for block in _docx_blocks(cell, f'{prefix}table:{tables}/row:{index}/cell:{column}/'):
                            # Cell paragraphs are already included by cell.text above.
                            if '/table:' in block['locator'][len(prefix):]:
                                yield block


def _parse_docx(path: Path) -> dict:
    doc = Document(str(path))
    blocks = list(_docx_blocks(doc))
    warnings = []
    with ZipFile(path) as archive:
        names = archive.namelist()
        if any(name.startswith('word/media/') for name in names):
            warnings.append('Word 内嵌图像未执行 OCR，图片文字须补充核对。')
        for name in names:
            if name.startswith(('word/header', 'word/footer', 'word/footnotes', 'word/endnotes')) and name.endswith('.xml'):
                # These parts may contain substantive requirements outside body order.
                from lxml import etree
                root = etree.fromstring(archive.read(name), parser=etree.XMLParser(resolve_entities=False))
                if root.xpath('//*[local-name()="t"]/text()'):
                    warnings.append('Word 页眉、页脚或脚注含文本，正文提取未包含，须人工核对。')
                    break
    xml = doc.element.xml
    if any(marker in xml for marker in ('<w:txbxContent', '<w:ins ', '<w:del ', '<w:sdt', '<w:altChunk')):
        warnings.append('Word 包含文本框、修订或内容控件，提取覆盖不完整，须人工核对。')
    if not blocks:
        raise DocumentError('Word contains no extractable text; image-only Word requires conversion and OCR')
    return {'blocks': blocks, 'warnings': warnings, 'complete': not warnings}


def parse_document(path: Path, *, converter: str = '', ocr_command: str | list[str] = '',
                   converter_remote: str = '', ocr_remote: str = '', remote_token: str = '') -> dict:
    """Extract actual PDF page text or DOCX body blocks, preserving source order.

    ``converter`` is an absolute LibreOffice executable. ``ocr_command`` is a
    local executable receiving INPUT_PDF OUTPUT_JSON; see docs for its schema.
    Missing capability, unsupported format, corruption and empty text raise.
    """
    path = Path(path)
    if not path.is_file():
        raise DocumentError('Document file does not exist')
    suffix = path.suffix.lower()
    try:
        if suffix == '.pdf':
            return _parse_pdf(path, ocr_command, ocr_remote, remote_token)
        if suffix == '.docx':
            return _parse_docx(path)
        if suffix == '.doc':
            if converter_remote:
                result = _remote_convert(path, converter_remote, remote_token)
            elif converter:
                with tempfile.TemporaryDirectory(prefix='bid-doc-') as temporary:
                    directory = Path(temporary)
                    profile = (directory / 'profile').as_uri()
                    _execute(converter, [f'-env:UserInstallation={profile}', '--headless',
                                         '--convert-to', 'docx', '--outdir', str(directory),
                                         str(path.resolve())], 'DOC converter')
                    converted = directory / f'{path.stem}.docx'
                    if not converted.is_file():
                        raise DocumentError('DOC converter produced no DOCX file')
                    result = _parse_docx(converted)
            else:
                raise DocumentError('Legacy DOC requires a configured local LibreOffice converter or doc-service')
            result['warnings'].append('旧版 DOC 已转换；定位指向转换后的段落和表格，不能作为原文件页码。')
            return result
        raise DocumentError(f'Unsupported document type: {suffix}')
    except DocumentError:
        raise
    except Exception as exc:
        raise DocumentError(f'Unable to parse {suffix} document: {type(exc).__name__}') from exc


def _field(paragraph, instruction: str, placeholder: str) -> None:
    field = OxmlElement('w:fldSimple')
    field.set(qn('w:instr'), instruction)
    field.set(qn('w:dirty'), 'true')
    run, text = OxmlElement('w:r'), OxmlElement('w:t')
    text.text = placeholder
    run.append(text)
    field.append(run)
    paragraph._p.append(field)


def _bookmark(paragraph, identifier: int) -> str:
    name = f'evidence_{identifier}'
    start, end = OxmlElement('w:bookmarkStart'), OxmlElement('w:bookmarkEnd')
    start.set(qn('w:id'), str(identifier))
    start.set(qn('w:name'), name)
    end.set(qn('w:id'), str(identifier))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)
    return name


def _caption_text(kind: str, item: dict, number: int, template: dict) -> str:
    """题注文本：只在条目自带 caption/title 时才生成（避免给历史产物凭空加题注）。"""
    title = str(item.get('caption') or item.get('title') or '').strip()
    if not title:
        return ''
    prefix = template['table']['caption_prefix'] if kind == 'table' else template['image']['caption_prefix']
    return f'{prefix} {number} {title}' if prefix else f'{number} {title}'


def _figure_reference(figure) -> str:
    """把章节配图的引用形式（字符串 / {file|name|path}）统一成字符串。"""
    if isinstance(figure, str):
        return figure.strip()
    if not isinstance(figure, dict):
        return ''
    return str(figure.get('file') or figure.get('name') or figure.get('path') or '').strip()


def _add_table(doc, table: dict, template: dict | None = None, caption: str = '') -> None:
    headers, rows = table.get('headers', []), table.get('rows', [])
    if not headers:
        raise DocumentError('Table requires nonempty headers')
    if any(len(row) != len(headers) for row in rows):
        raise DocumentError('Table row width does not match headers')
    if caption and template:
        bid_template.add_caption(doc, caption, size_pt=template['table']['font_size_pt'])
    result = doc.add_table(rows=1, cols=len(headers))
    result.style = template['table']['style'] if template else 'Table Grid'
    for cell, value in zip(result.rows[0].cells, headers):
        cell.text = str(value)
    properties = result.rows[0]._tr.get_or_add_trPr()
    properties.append(OxmlElement('w:tblHeader'))
    for row in rows:
        for cell, value in zip(result.add_row().cells, row):
            cell.text = str(value)
    if template:
        bid_template.style_table(result, template)


def _attach(doc, path: Path, max_width_cm: float = 16.0, max_height_cm: float = 22.0) -> None:
    def insert(image):
        image = ImageOps.exif_transpose(image).convert('RGB')
        width, height = image.size
        scale = min(max_width_cm / width, max_height_cm / height)
        stream = io.BytesIO()
        image.save(stream, format='PNG')
        stream.seek(0)
        doc.add_picture(stream, width=Cm(width * scale), height=Cm(height * scale))

    if path.suffix.lower() == '.pdf':
        with pdfium.PdfDocument(str(path)) as pdf:
            if not len(pdf):
                raise DocumentError('PDF attachment contains no pages')
            for index in range(len(pdf)):
                if index:
                    doc.add_page_break()
                page = pdf[index]
                try:
                    # Bound each rendered page even for oversized PDF page boxes.
                    scale = min(1.5, 2000 / max(page.get_size()))
                    bitmap = page.render(scale=scale)
                    try:
                        with bitmap.to_pil() as image:
                            insert(image)
                    finally:
                        bitmap.close()
                finally:
                    page.close()
    else:
        with Image.open(path) as image:
            insert(image)


def build_bid_documents(output_dir: Path, project: dict, sections: list[dict],
                        attachments: list[dict], quote: dict | None, *,
                        final: bool = False, missing: list[str] | None = None,
                        template=None) -> dict:
    """Build two DOCX volumes from supplied facts; no content is invented.

    The caller owns rule verification and must pass unresolved requirements in
    ``missing``. This worker cannot certify facts or Word pagination.

    ``template`` 可为预设名 / 用户模板名 / 内联模板对象；为空时用 ``standard``，
    其排版参数与历史硬编码输出一致（有回归测试保证）。
    """
    missing = missing or []
    if final and missing:
        raise DocumentError('Final output blocked by missing requirements: ' + '; '.join(missing))
    for item in [*sections, *attachments]:
        if item.get('volume') not in VOLUMES:
            raise DocumentError('Each section and attachment requires a valid volume')
        if item['volume'] == 'business_technical' and item.get('contains_current_quote'):
            raise DocumentError('Current quote content cannot enter business_technical volume')
    for attachment in attachments:
        path = Path(attachment.get('path', ''))
        if path.suffix.lower() not in {'.png', '.jpg', '.jpeg', '.pdf'} or not path.is_file():
            raise DocumentError('Unsupported or missing attachment; only local PNG/JPEG/PDF accepted')
    resolved = bid_template.resolve_template(template)
    warnings = ['Word 目录、附件页码及分页尚未渲染校验；须在 Word 更新全部域并核对版式。']
    if missing:
        warnings.append('草稿存在待补项：' + '；'.join(missing))
    if resolved['name'] != 'standard':
        warnings.append("已套用导出模板「{}」（{} / {}）。".format(
            resolved['label'], resolved['page']['paper'], resolved['page']['orientation']))
    output_dir = Path(output_dir)
    missing_figures: list[str] = []
    html_ignored: list[str] = []
    # Build both in memory first: malformed proof or table never publishes a partial pair.
    documents = {}
    try:
        for volume, label in VOLUMES.items():
            doc = Document()
            bid_template.apply_document_template(
                doc, resolved,
                banner='' if final else '草稿 DRAFT  待核验 禁止直接提交')
            title = project.get('name') or project.get('project_name') or '投标文件'
            doc.add_paragraph(str(title), style='Title')
            doc.add_paragraph(label, style='Subtitle')
            doc.add_paragraph('状态：定稿' if final else '状态：草稿 待核验')
            if project.get('bid_no'):
                doc.add_paragraph('项目编号：' + str(project['bid_no']))
            if missing:
                doc.add_paragraph('待补项：' + '；'.join(missing))
            doc.add_page_break()
            doc.add_paragraph('目录', style='Title')
            _field(doc.add_paragraph(), 'TOC \\o "1-3" \\h \\z \\u', '请在 Word 中更新目录')
            doc.add_page_break()
            table_index = figure_index = 0
            for item in sections:
                if item['volume'] != volume:
                    continue
                doc.add_heading(str(item.get('title', '')), level=1)
                if item.get('html'):
                    # 受限 HTML 正文：结构化内容交给程序确定性转 Word；校验不过直接报错不静默降级
                    if (str(item.get('content') or '').strip() or item.get('table') or item.get('images')):
                        html_ignored.append(str(item.get('title') or '')[:40])
                    rendered = bid_html.render_html(
                        doc, item['html'], resolved, figure_number=figure_index,
                        table_number=table_index, missing=missing_figures)
                    figure_index, table_index = rendered['figure_number'], rendered['table_number']
                    continue
                for line in str(item.get('content', '')).splitlines():
                    doc.add_paragraph(line)
                if item.get('table'):
                    caption = _caption_text('table', item['table'], table_index + 1, resolved)
                    if caption:
                        table_index += 1
                    _add_table(doc, item['table'], resolved, caption)
                for figure in item.get('images') or []:
                    raw = _figure_reference(figure)
                    path = bid_html.resolve_figure_path(raw)
                    if path is None:
                        missing_figures.append('配图未找到：' + raw[:80])
                        continue
                    caption = _caption_text('image', figure if isinstance(figure, dict) else {},
                                            figure_index + 1, resolved)
                    if caption:
                        figure_index += 1
                    bid_template.add_figure(doc, path, resolved, caption=caption)
            if volume == 'price' and quote:
                doc.add_heading(str(quote.get('title', '报价表')), level=1)
                _add_table(doc, quote, resolved)
                for note in quote.get('notes', []):
                    doc.add_paragraph(str(note))
            proofs = [item for item in attachments if item['volume'] == volume]
            if proofs:
                doc.add_heading('附件索引', level=1)
                for identifier, attachment in enumerate(proofs, 1):
                    paragraph = doc.add_paragraph(f'{identifier}. {attachment.get("title", "证明材料")}  页码：')
                    _field(paragraph, f'PAGEREF evidence_{identifier} \\h', '待更新')
                for identifier, attachment in enumerate(proofs, 1):
                    doc.add_page_break()
                    heading = doc.add_heading(str(attachment.get('title', '证明材料')), level=1)
                    _bookmark(heading, identifier)
                    _attach(doc, Path(attachment['path']),
                            resolved['image']['max_width_cm'], resolved['image']['max_height_cm'])
            settings = doc.settings.element
            update = OxmlElement('w:updateFields')
            update.set(qn('w:val'), 'true')
            settings.append(update)
            buffer = io.BytesIO()
            doc.save(buffer)
            documents[volume] = buffer.getvalue()
        output_dir.mkdir(parents=True, exist_ok=True)
        files = []
        for volume, data in documents.items():
            path = output_dir / f'{volume}.docx'
            with tempfile.NamedTemporaryFile(dir=output_dir, suffix='.docx', delete=False) as temporary:
                temporary.write(data)
                temporary_path = Path(temporary.name)
            try:
                os.replace(temporary_path, path)
            finally:
                temporary_path.unlink(missing_ok=True)
            files.append({'name': path.name, 'path': str(path.resolve()), 'kind': volume})
        if missing_figures:
            warnings.append('以下配图未找到已跳过（不影响其余内容）：' + '；'.join(missing_figures))
        if html_ignored:
            warnings.append('以下小节同时提供了 html 与 content/table/images，已按 html 渲染并忽略其余字段：'
                            + '、'.join(html_ignored))
        return {'files': files, 'warnings': warnings}
    except (DocumentError, bid_html.HtmlError):
        # HtmlError 带着精确的校验定位信息，不能被下面的通用包装吞掉
        raise
    except Exception as exc:
        raise DocumentError(f'Unable to build document: {type(exc).__name__}') from exc
