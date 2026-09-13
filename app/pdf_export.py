"""Генерация PDF-отчёта. Кроссплатформенная поддержка кириллицы.

Логика поиска шрифта:
    1. DejaVuSans (Linux: /usr/share/fonts)
    2. Arial (Windows: C:/Windows/Fonts)
    3. Helvetica (встроен в reportlab — fallback, без кириллицы)

Если ни один кириллический шрифт не найден, PDF будет с латиницей
(цифры и коды ошибок отобразятся, текст — нет). В README есть
инструкция как установить DejaVuSans на Windows.
"""
from __future__ import annotations

import io
import os
from datetime import datetime
from typing import Any

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# ─── Поиск шрифта ───

_FONT = "Helvetica"  # fallback
_FONT_B = "Helvetica-Bold"
_REGISTERED = False

_CANDIDATES = [
    # Linux
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
     "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
     "DejaVuSans", "DejaVuSans-Bold"),
    ("/usr/share/fonts/truetype/freefont/FreeSans.ttf",
     "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
     "FreeSans", "FreeSans-Bold"),
    # Windows
    ("C:/Windows/Fonts/arial.ttf",
     "C:/Windows/Fonts/arialbd.ttf",
     "ArialCyr", "ArialCyr-Bold"),
    ("C:\\Windows\\Fonts\\arial.ttf",
     "C:\\Windows\\Fonts\\arialbd.ttf",
     "ArialCyr2", "ArialCyr2-Bold"),
    # Mac
    ("/System/Library/Fonts/Supplemental/Arial.ttf",
     "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
     "ArialMac", "ArialMac-Bold"),
]


def _ensure_fonts():
    global _FONT, _FONT_B, _REGISTERED
    if _REGISTERED:
        return
    _REGISTERED = True

    for regular, bold, name, name_b in _CANDIDATES:
        if os.path.exists(regular):
            try:
                pdfmetrics.registerFont(TTFont(name, regular))
                _FONT = name
                if os.path.exists(bold):
                    pdfmetrics.registerFont(TTFont(name_b, bold))
                    _FONT_B = name_b
                else:
                    _FONT_B = name  # bold не найден — используем regular
                return
            except Exception:
                continue


SEVERITY_COLORS = {
    "high": colors.HexColor("#c04040"),
    "medium": colors.HexColor("#c47a1e"),
    "low": colors.HexColor("#3a6ea5"),
}
SEVERITY_LABELS = {"high": "Критичное", "medium": "Среднее", "low": "Незначительное"}


# Подпись вида проверки в шапке отчёта: язык кода или нормоконтроль .docx
_CHECK_LABELS = {
    "python": "Python (PEP 8)",
    "javascript": "JavaScript (стиль)",
    "sql": "SQL (стиль)",
    "java": "Java (конвенции)",
    "cpp": "C/C++ (стиль)",
}


def _check_label(report: dict) -> str:
    return _CHECK_LABELS.get(report.get("file_type"), "Нормоконтроль ГОСТ/АГУ")


def generate_pdf_report(report: dict[str, Any]) -> bytes:
    _ensure_fonts()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)

    title_s = ParagraphStyle("T", fontName=_FONT_B, fontSize=16, spaceAfter=8, leading=20)
    sub_s = ParagraphStyle("S", fontName=_FONT, fontSize=9, textColor=colors.HexColor("#666"), spaceAfter=16, leading=13)
    head_s = ParagraphStyle("H", fontName=_FONT_B, fontSize=12, spaceBefore=16, spaceAfter=10, leading=16)
    issue_s = ParagraphStyle("I", fontName=_FONT, fontSize=8.5, leading=12, spaceAfter=2)
    footer_s = ParagraphStyle("F", fontName=_FONT, fontSize=7, textColor=colors.HexColor("#999"))
    th_s = ParagraphStyle("TH", fontName=_FONT, fontSize=8, textColor=colors.HexColor("#666"), alignment=1)
    tv_s = ParagraphStyle("TV", fontName=_FONT_B, fontSize=14, alignment=1)

    elems = []

    elems.append(Paragraph("SmartDoc &amp; Code Review", title_s))

    label = _check_label(report)
    now = datetime.now().strftime("%d.%m.%Y %H:%M")
    elems.append(Paragraph(
        f"Файл: <b>{_s(report['filename'])}</b> | Проверка: {label} | Дата: {now}", sub_s))
    elems.append(Spacer(1, 2*mm))

    # Сводка
    elems.append(Paragraph("Сводная статистика", head_s))
    sm = report.get("summary", {})
    data = [
        [Paragraph("Всего", th_s), Paragraph("Критичных", th_s),
         Paragraph("Средних", th_s), Paragraph("Незначительных", th_s)],
        [Paragraph(str(report.get("total_issues", 0)), tv_s),
         Paragraph(f'<font color="#c04040">{sm.get("high", 0)}</font>', tv_s),
         Paragraph(f'<font color="#c47a1e">{sm.get("medium", 0)}</font>', tv_s),
         Paragraph(f'<font color="#3a6ea5">{sm.get("low", 0)}</font>', tv_s)],
    ]
    t = Table(data, colWidths=[3.8*cm]*4)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f5f3ee")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d8d4cb")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elems.append(t)
    elems.append(Spacer(1, 6*mm))

    # Замечания
    issues = report.get("issues", [])
    source_lines = report.get("source_lines", [])

    if not issues:
        elems.append(Paragraph("Замечаний не найдено. Файл соответствует требованиям.", issue_s))
    else:
        elems.append(Paragraph(f"Подробный список ({len(issues)} замечаний)", head_s))

        # Стиль для фрагментов кода
        code_s = ParagraphStyle("CODE", fontName=_FONT, fontSize=7.5, leading=10,
                                textColor=colors.HexColor("#444"),
                                backColor=colors.HexColor("#f5f3ee"),
                                borderPadding=(3, 6, 3, 6),
                                spaceAfter=3)

        for idx, iss in enumerate(issues, 1):
            sev = iss.get("severity", "medium")
            col = SEVERITY_COLORS.get(sev, colors.grey).hexval()
            lab = SEVERITY_LABELS.get(sev, sev)
            if "line" in iss:
                loc = f"стр. {iss['line']}"
            else:
                loc = _s((iss.get("location") or "")[:70])
            desc = _s(iss.get("description") or iss.get("message", ""))
            code = _s(iss.get("code", ""))

            txt = f'<font color="{col}">[{_s(lab)}]</font> <b>{code}</b> — {desc}'
            if loc:
                txt += f'  <font color="#999">({loc})</font>'
            if iss.get("expected") and iss.get("actual"):
                txt += f'<br/><font color="#999" size="7">  Требуется: {_s(iss["expected"])} | Фактически: {_s(iss["actual"])}</font>'
            elems.append(Paragraph(txt, issue_s))

            # Фрагмент исходного кода (±1 строка контекста)
            line_no = iss.get("line")
            if line_no and source_lines and 1 <= line_no <= len(source_lines):
                start = max(0, line_no - 2)
                end = min(len(source_lines), line_no + 1)
                snippet_parts = []
                for li in range(start, end):
                    ln = li + 1
                    prefix = "▸ " if ln == line_no else "  "
                    snippet_parts.append(f"{prefix}{ln:>4} │ {_s(source_lines[li])}")
                snippet = "<br/>".join(snippet_parts)
                elems.append(Paragraph(snippet, code_s))

            elems.append(Spacer(1, 1.5*mm))
            if idx >= 200:
                elems.append(Paragraph(f"... и ещё {len(issues) - 200} замечаний", issue_s))
                break

    elems.append(Spacer(1, 12*mm))
    elems.append(Paragraph(f"SmartDoc &amp; Code Review · Данил Полстянов · {now}", footer_s))

    doc.build(elems)
    return buf.getvalue()


def _s(text: str) -> str:
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
