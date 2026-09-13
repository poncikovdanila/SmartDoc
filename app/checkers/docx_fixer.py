"""Модуль автоисправления .docx по требованиям нормоконтроля АГУ (ФЦТиК).

Принципы (v8, «неразрушающее исправление»):

    1. Содержательный текст НЕ меняется. Исключения: точка в конце заголовка,
       двойные пробелы и пробел перед знаком препинания в обычных абзацах
       (в заголовках и оглавлении пробелы не трогаются — там они бывают
       намеренными).
    2. Абзацы НЕ удаляются и НЕ добавляются, кроме пустых абзацев между
       рисунком и его подписью и повторных пустых строк в основном тексте.
       Абзац, несущий разрыв раздела (<w:sectPr>), закладку или рисунок,
       не удаляется никогда.
    3. Титульный лист и всё до начала основного текста (``_find_body_start``)
       не трогается вообще, кроме полей страницы.
    4. Исправляется только то, что отклоняется от правил. Если абзац уже
       оформлен верно (через стиль или напрямую) — он остаётся как есть.
       Явное форматирование run-ов не «размазывается» по всему документу:
       шрифт и кегль задаются на уровне стиля Normal, а на run — только там,
       где run явно задаёт неверное значение.
    5. Центрированные и выровненные вправо абзацы основного текста, элементы
       списков (с нумерацией Word), оглавление, подписи и абзацы с рисунками
       оформляются по своим правилам, а не как обычный текст.

Что исправляется:
    * поля страницы → 3.5/1/2.5/2.5 см (во всех разделах)
    * стиль Normal → Times New Roman 14 пт, чёрный
    * стили заголовков → Times New Roman, полужирный, чёрный
    * основной текст: выравнивание по ширине, отступ 1.25 см,
      межстрочный 1.5, интервалы до/после > 6 пт → 0
    * run-ы с явно неверным шрифтом/кеглем/цветом → исправляются точечно
    * заголовки: явное bold=False снимается, точка в конце убирается
    * подписи к рисункам → по центру без отступа;
      подписи к таблицам → влево без отступа
    * рисунки → по центру без отступа
    * таблицы: шрифт/цвет в ячейках точечно, кегль > 14 → 14
    * пустые абзацы между рисунком и подписью → удаляются (с защитой)
"""
from __future__ import annotations

import io
import re
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from app.checkers.docx_checker import (
    ACCEPTED_FONT_SIZES,
    TOLERANCE_CM,
    TOLERANCE_SPACING,
    _find_body_start,
    _get_rules,
    _is_text_paragraph,
    _resolve_first_line_indent_cm,
    _resolve_font_name,
    _resolve_font_size_pt,
    _resolve_line_spacing,
)

BLACK = RGBColor(0, 0, 0)
DOUBLE_SPACE_RE = re.compile(r" {2,}")
SPACE_BEFORE_PUNCT_RE = re.compile(r" +([,;!?]|\.(?=\s|$))")
SPACING_THRESHOLD = Pt(6)  # синхронно с check_paragraph_spacing

FIGURE_CAPTION_RE = re.compile(r"^\s*(Рис(\.|унок)|Figure)\s*[\dА-ЯA-Z]", re.IGNORECASE)
TABLE_CAPTION_RE = re.compile(
    r"^\s*(Таблица|Table|Продолжение\s+таблицы|Окончание\s+таблицы)\s*[\dА-ЯA-Z]",
    re.IGNORECASE,
)


# ───────────────────────── классификация абзацев ─────────────────────────

def _paragraph_has_image(paragraph) -> bool:
    xml = paragraph._element.xml
    return "blipFill" in xml or "w:drawing" in xml or "w:pict" in xml


def _has_section_break(paragraph) -> bool:
    ppr = paragraph._element.pPr
    return ppr is not None and ppr.find(qn("w:sectPr")) is not None


def _has_bookmark(paragraph) -> bool:
    return paragraph._element.find(qn("w:bookmarkStart")) is not None


def _is_list_item(paragraph) -> bool:
    """Абзац с нумерацией/маркерами Word (numPr) или стилем списка."""
    ppr = paragraph._element.pPr
    if ppr is not None and ppr.find(qn("w:numPr")) is not None:
        return True
    style_name = (paragraph.style.name or "").lower()
    return "list" in style_name or "список" in style_name


def _is_figure_caption(paragraph) -> bool:
    return bool(FIGURE_CAPTION_RE.match(paragraph.text))


def _is_table_caption(paragraph) -> bool:
    return bool(TABLE_CAPTION_RE.match(paragraph.text))


def _is_caption(paragraph) -> bool:
    return _is_figure_caption(paragraph) or _is_table_caption(paragraph)


def _is_heading(paragraph) -> bool:
    style_name = (paragraph.style.name or "").lower()
    if "toc" in style_name:
        return False
    return ("heading" in style_name or "заголов" in style_name
            or style_name.startswith("+") or "раздел" in style_name)


def _is_toc(paragraph) -> bool:
    return "toc" in (paragraph.style.name or "").lower()


def _is_empty(paragraph) -> bool:
    return not paragraph.text.strip() and not _paragraph_has_image(paragraph)


def _has_break(paragraph) -> bool:
    """Абзац содержит разрыв страницы/колонки/строки (<w:br>) или pageBreakBefore."""
    xml = paragraph._element.xml
    return "<w:br" in xml or "w:pageBreakBefore" in xml


def _is_protected(paragraph) -> bool:
    """Абзац, который нельзя удалять ни при каких условиях."""
    return (_has_section_break(paragraph) or _has_bookmark(paragraph)
            or _paragraph_has_image(paragraph) or _has_break(paragraph))


# ───────────────────────── низкоуровневые правки ─────────────────────────

def _set_font_name_everywhere(font_obj, rpr_parent, font_name: str) -> None:
    """Выставляет имя шрифта во всех четырёх атрибутах rFonts."""
    font_obj.name = font_name
    rpr = rpr_parent.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), font_name)
    # Тема оформления может перебивать явное имя — убираем ссылки на тему
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]


def _run_explicit_font_name(run) -> str | None:
    rpr = run._element.rPr
    if rpr is None:
        return None
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        return None
    return rfonts.get(qn("w:ascii")) or rfonts.get(qn("w:hAnsi"))


def _run_has_color(run) -> bool:
    color = run.font.color
    return color is not None and color.type is not None


def _fix_run(run, font_name: str, size_pt: float, *, force_size: bool = False) -> None:
    """Точечно исправляет только явно неверные свойства run-а."""
    if not run.text.strip():
        return
    explicit = _run_explicit_font_name(run)
    if explicit is not None and explicit != font_name:
        _set_font_name_everywhere(run.font, run._element, font_name)
    size = run.font.size
    if size is not None:
        if force_size:
            if abs(size.pt - size_pt) > 0.5:
                run.font.size = Pt(size_pt)
        elif size.pt not in _accepted_sizes(size_pt):
            run.font.size = Pt(size_pt)
    if _run_has_color(run):
        rgb = run.font.color.rgb
        if run.font.color.type is not None and (rgb is None or rgb != BLACK):
            run.font.color.rgb = BLACK
    if run.font.highlight_color is not None:
        run.font.highlight_color = None


def _accepted_sizes(size_pt: float) -> set[float]:
    """Для стандартных правил АГУ допустимы 12 и 14 пт (эталон кафедры набран
    12 пт); для кастомных правил — только заданный кегль."""
    return set(ACCEPTED_FONT_SIZES) if abs(size_pt - 14.0) < 0.5 else {float(size_pt)}


def _fix_style_font(style, font_name: str, size_pt: float | None,
                    *, bold: bool | None = None) -> None:
    font = style.font
    if font.name != font_name:
        _set_font_name_everywhere(font, style.element, font_name)
    if size_pt is not None and (font.size is None or font.size.pt not in _accepted_sizes(size_pt)):
        font.size = Pt(size_pt)
    if bold is not None and font.bold is not True:
        font.bold = bold
    if font.color is not None and font.color.type is not None:
        if font.color.rgb is None or font.color.rgb != BLACK:
            font.color.rgb = BLACK


def _style_chain_bold(paragraph) -> bool:
    style = paragraph.style
    while style is not None:
        if style.font.bold:
            return True
        style = style.base_style
    return False


def _resolve_alignment(paragraph):
    """Выравнивание с учётом наследования от стиля (как его увидит Word)."""
    if paragraph.alignment is not None:
        return paragraph.alignment
    style = paragraph.style
    while style is not None:
        al = style.paragraph_format.alignment
        if al is not None:
            return al
        style = style.base_style
    return WD_ALIGN_PARAGRAPH.LEFT


def _remove_paragraph(paragraph) -> None:
    elem = paragraph._element
    elem.getparent().remove(elem)


# ───────────────────────── основная функция ─────────────────────────

def autofix_docx(file_path: Path, custom_rules: dict | None = None) -> bytes:
    """Открывает .docx, исправляет нарушения оформления, возвращает байты."""
    rules = _get_rules(custom_rules)
    r_font: str = rules["font_name"]
    r_size: float = rules["font_size_pt"]
    r_spacing: float = rules["line_spacing"]
    r_indent_cm: float = rules["first_line_indent_cm"]
    r_margins = {k: Cm(v) for k, v in rules["margins_cm"].items()}

    document = Document(file_path)

    # 1. Поля страницы — во всех разделах, включая титульный
    for section in document.sections:
        for side in ("left", "right", "top", "bottom"):
            current = getattr(section, f"{side}_margin")
            if current is None or abs(current.cm - rules["margins_cm"][side]) > TOLERANCE_CM:
                setattr(section, f"{side}_margin", r_margins[side])

    # 2. Стили: Normal и заголовки. Через стиль исправляется 90 % документа,
    #    не засоряя run-ы явным форматированием.
    try:
        _fix_style_font(document.styles["Normal"], r_font, r_size)
    except KeyError:
        pass
    for style in document.styles:
        name = (style.name or "").lower()
        if style.type != WD_STYLE_TYPE.PARAGRAPH:
            continue
        if "toc" in name:
            continue
        if "heading" in name or "заголов" in name or name.startswith("+") or "раздел" in name:
            _fix_style_font(style, r_font, None, bold=True)

    # 3. Основной текст — только после титульного листа
    body_start = _find_body_start(document)
    paragraphs = document.paragraphs

    # 3a. Пустые абзацы между рисунком и подписью — удаляем (с защитой)
    to_remove: list = []
    for i in range(body_start, len(paragraphs)):
        if not _paragraph_has_image(paragraphs[i]):
            continue
        j = i + 1
        blanks = []
        while j < len(paragraphs) and _is_empty(paragraphs[j]) and not _is_protected(paragraphs[j]):
            blanks.append(paragraphs[j])
            j += 1
        if j < len(paragraphs) and _is_figure_caption(paragraphs[j]):
            to_remove.extend(blanks)
    # 3b. Повторные пустые строки в основном тексте: 2+ подряд → 1
    consecutive = 0
    for i in range(body_start, len(paragraphs)):
        p = paragraphs[i]
        if _is_empty(p):
            consecutive += 1
            if consecutive >= 2 and not _is_protected(p) and p not in to_remove:
                to_remove.append(p)
        else:
            consecutive = 0
    for p in to_remove:
        _remove_paragraph(p)

    # 3c. Проход по абзацам основного текста
    paragraphs = document.paragraphs
    body_start = _find_body_start(document)
    for idx in range(body_start, len(paragraphs)):
        paragraph = paragraphs[idx]
        pf = paragraph.paragraph_format

        if _is_toc(paragraph) or _is_empty(paragraph):
            continue

        # Рисунок (абзац без текста) — по центру, без отступа
        if _paragraph_has_image(paragraph) and not paragraph.text.strip():
            if _resolve_alignment(paragraph) != WD_ALIGN_PARAGRAPH.CENTER:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if _resolve_first_line_indent_cm(paragraph, document) > TOLERANCE_CM:
                pf.first_line_indent = Cm(0)
            continue

        # Подписи
        if _is_caption(paragraph):
            current = _resolve_alignment(paragraph)
            if _is_figure_caption(paragraph):
                if current != WD_ALIGN_PARAGRAPH.CENTER:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif current in (WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT):
                # Подпись таблицы — слева (по ширине для одной строки равнозначно)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            if abs(_resolve_first_line_indent_cm(paragraph, document)) > TOLERANCE_CM:
                pf.first_line_indent = Cm(0)
            if abs((_resolve_line_spacing(paragraph, document) or 0) - r_spacing) > TOLERANCE_SPACING:
                pf.line_spacing = r_spacing
            for run in paragraph.runs:
                _fix_run(run, r_font, r_size)
            continue

        # Заголовки
        if _is_heading(paragraph):
            style_bold = _style_chain_bold(paragraph)
            for run in paragraph.runs:
                if run.bold is False:
                    run.bold = None if style_bold else True
                elif not style_bold and run.bold is None and run.text.strip():
                    run.bold = True
                _fix_run(run, r_font, r_size)
            # Точка в конце заголовка — единственная правка текста
            if paragraph.text.rstrip().endswith("."):
                for run in reversed(paragraph.runs):
                    if run.text.rstrip().endswith("."):
                        run.text = run.text.rstrip()[:-1]
                        break
            continue

        # Элементы списков: отступы (в т.ч. висячие) не трогаем
        if _is_list_item(paragraph):
            if abs((_resolve_line_spacing(paragraph, document) or 0) - r_spacing) > TOLERANCE_SPACING:
                pf.line_spacing = r_spacing
            for run in paragraph.runs:
                _fix_run(run, r_font, r_size)
            continue

        if not _is_text_paragraph(paragraph):
            continue

        # Обычный текст
        alignment = _resolve_alignment(paragraph)
        is_centered = alignment in (WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT)
        if not is_centered:
            if alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            indent = _resolve_first_line_indent_cm(paragraph, document)
            if indent is None or abs(indent - r_indent_cm) > TOLERANCE_CM:
                pf.first_line_indent = Cm(r_indent_cm)
                if pf.left_indent is not None and pf.left_indent.cm < 0:
                    pf.left_indent = Cm(0)
        spacing = _resolve_line_spacing(paragraph, document)
        if spacing is None or abs(spacing - r_spacing) > TOLERANCE_SPACING:
            pf.line_spacing = r_spacing
        if pf.space_before is not None and pf.space_before > SPACING_THRESHOLD:
            pf.space_before = Pt(0)
        if pf.space_after is not None and pf.space_after > SPACING_THRESHOLD:
            pf.space_after = Pt(0)

        # Шрифт: если не разрешается ни через run, ни через стиль — задаём явно
        if _resolve_font_name(paragraph, document) != r_font:
            for run in paragraph.runs:
                if run.text.strip():
                    _set_font_name_everywhere(run.font, run._element, r_font)
        size = _resolve_font_size_pt(paragraph, document)
        if size is None or size not in _accepted_sizes(r_size):
            for run in paragraph.runs:
                if run.text.strip() and (run.font.size is None
                                         or run.font.size.pt not in _accepted_sizes(r_size)):
                    run.font.size = Pt(r_size)
        for run in paragraph.runs:
            _fix_run(run, r_font, r_size)
            # Двойные пробелы и пробел перед знаком препинания — только в
            # обычном тексте (в заголовках/оглавлении пробелы могут быть намеренными)
            cleaned = DOUBLE_SPACE_RE.sub(" ", run.text)
            cleaned = SPACE_BEFORE_PUNCT_RE.sub(r"\1", cleaned)
            if cleaned != run.text:
                run.text = cleaned

    # 4. Таблицы: точечно шрифт/цвет, кегль > 14 → 14
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        _fix_run(run, r_font, r_size)
                        if run.font.size is not None and run.font.size.pt > 14:
                            run.font.size = Pt(r_size)

    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()
