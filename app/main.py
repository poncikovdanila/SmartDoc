"""SmartDoc & Code Review v8.

Маршруты:
    GET  /               — главная
    POST /api/check      — проверка (код или .docx) → JSON-отчёт
    POST /api/check-batch — пакетная проверка нескольких файлов
    POST /api/autofix    — автоисправление → скачивание файла
    POST /api/autofix-preview — предпросмотр автоисправления
    POST /api/export-pdf — экспорт отчёта в PDF
    POST /api/generate-template — шаблон .docx по текущим правилам
    GET  /api/presets    — список пресетов правил для .docx
    GET  /api/languages  — поддерживаемые языки и форматы
    GET  /health         — health-check

Поддерживаемые форматы: .docx (нормоконтроль АГУ/ГОСТ) и код на Python,
JavaScript, SQL, Java, C/C++. Автоисправление есть для .py и .docx.
"""
from __future__ import annotations

import os
import tempfile
import uuid
from pathlib import Path
from typing import Annotated

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from urllib.parse import quote

from app.checkers.code_checker import check_python_code
from app.checkers.code_fixer import autofix_python_code
from app.checkers.docx_checker import check_docx_document, PRESETS
from app.checkers.docx_fixer import autofix_docx
from app.checkers.multi_lang_checker import (
    AUTOFIX_EXTENSIONS,
    LANGUAGE_NAMES,
    SUPPORTED_CODE_EXTENSIONS,
    check_code_file,
)
from app.pdf_export import generate_pdf_report
from app.template_generator import generate_template

BASE_DIR = Path(__file__).resolve().parent
MAX_FILE_SIZE = 5 * 1024 * 1024
ALLOWED_EXTENSIONS = {".docx"} | SUPPORTED_CODE_EXTENSIONS
# Порядок для сообщений об ошибках: сначала самые частые форматы
_EXT_ORDER = [".docx", ".py", ".js", ".sql", ".java", ".cpp", ".c", ".h"]
ALLOWED_EXTENSIONS_TEXT = ", ".join(
    [e for e in _EXT_ORDER if e in ALLOWED_EXTENSIONS]
    + sorted(ALLOWED_EXTENSIONS - set(_EXT_ORDER))
)


def _content_disposition(filename: str) -> str:
    """Формирует заголовок Content-Disposition с поддержкой кириллицы (RFC 5987)."""
    ascii_name = filename.encode("ascii", errors="replace").decode("ascii")
    encoded_name = quote(filename)
    return f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{encoded_name}"


app = FastAPI(
    title="SmartDoc & Code Review",
    description=(
        "Проверка кода (Python/PEP 8, JavaScript, SQL, Java, C/C++) "
        "и документов .docx (нормоконтроль ГОСТ/АГУ)."
    ),
    version="8.0.0",
)

app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
templates = Jinja2Templates(directory=BASE_DIR / "templates")


@app.get("/rules", include_in_schema=False)
async def rules_document():
    """Описание требований нормоконтроля АГУ ФЦТиК (markdown как текст)."""
    from fastapi.responses import PlainTextResponse
    path = BASE_DIR.parent / "docs" / "Требования_нормоконтроля_АГУ_ФЦТиК.md"
    if not path.exists():
        raise HTTPException(404, "Документ не найден")
    return PlainTextResponse(path.read_text(encoding="utf-8"), media_type="text/markdown; charset=utf-8")


def _parse_rules(docx_rules: str | None) -> dict | None:
    """Разбирает JSON с пользовательскими правилами .docx. Мусор игнорируем."""
    if not docx_rules:
        return None
    import json as _json
    try:
        parsed = _json.loads(docx_rules)
    except (ValueError, TypeError):
        return None
    return parsed if isinstance(parsed, dict) else None


def _run_check(tmp: Path, ext: str, filename: str, custom_rules: dict | None) -> dict:
    """Направляет файл в нужный чекер по расширению."""
    if ext == ".docx":
        return check_docx_document(tmp, filename, custom_rules)
    if ext == ".py":
        return check_python_code(tmp, filename)
    return check_code_file(tmp, filename, ext)


def _validate(file: UploadFile, content: bytes) -> str:
    if not file.filename:
        raise HTTPException(400, "Имя файла не указано")
    ext = Path(file.filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            400,
            f"Формат {ext} не поддерживается. "
            f"Допустимы: {ALLOWED_EXTENSIONS_TEXT}",
        )
    if len(content) > MAX_FILE_SIZE:
        raise HTTPException(413, f"Файл слишком большой (макс. {MAX_FILE_SIZE // 1024 // 1024} МБ)")
    if not content:
        raise HTTPException(400, "Файл пуст")
    return ext


def _require_autofix_support(ext: str) -> None:
    """Автоисправление есть не для всех форматов — остальным отвечаем понятно."""
    if ext in AUTOFIX_EXTENSIONS:
        return
    language = LANGUAGE_NAMES.get(ext, ext)
    raise HTTPException(
        422,
        f"Автоисправление для {language} пока не поддерживается — "
        f"доступна только проверка. Исправлять автоматически умеем .py и .docx",
    )


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(request, "index.html")


@app.post("/api/check")
async def check_file(file: Annotated[UploadFile, File(...)],
                     docx_rules: Annotated[str | None, Form()] = None):
    content = await file.read()
    ext = _validate(file, content)
    tmp = Path(tempfile.gettempdir()) / f"sd_{uuid.uuid4().hex}{ext}"

    # Пользовательские правила нужны только для .docx, но разбираем всегда
    custom_rules = _parse_rules(docx_rules)

    try:
        tmp.write_bytes(content)
        report = _run_check(tmp, ext, file.filename or "", custom_rules)
        return JSONResponse(content=report)
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


@app.post("/api/check-batch")
async def check_batch(files: Annotated[list[UploadFile], File(...)],
                      docx_rules: Annotated[str | None, Form()] = None):
    """Пакетная проверка нескольких файлов."""
    custom_rules = _parse_rules(docx_rules)

    reports = []
    total_issues = 0
    summary = {"high": 0, "medium": 0, "low": 0}

    for file in files:
        content = await file.read()
        try:
            ext = _validate(file, content)
        except HTTPException as e:
            reports.append({
                "filename": file.filename or "unknown",
                "file_type": "unknown",
                "total_issues": 0,
                "summary": {"high": 0, "medium": 0, "low": 0},
                "issues": [],
                "error": e.detail,
            })
            continue

        tmp = Path(tempfile.gettempdir()) / f"sd_b_{uuid.uuid4().hex}{ext}"
        try:
            tmp.write_bytes(content)
            report = _run_check(tmp, ext, file.filename or "", custom_rules)
            reports.append(report)
            total_issues += report["total_issues"]
            for sev in ("high", "medium", "low"):
                summary[sev] += report["summary"].get(sev, 0)
        except Exception as e:
            reports.append({
                "filename": file.filename or "unknown",
                "file_type": "unknown",
                "total_issues": 0,
                "summary": {"high": 0, "medium": 0, "low": 0},
                "issues": [],
                "error": str(e),
            })
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass

    return JSONResponse(content={
        "batch": True,
        "file_count": len(reports),
        "total_issues": total_issues,
        "summary": summary,
        "reports": reports,
    })


@app.post("/api/autofix-preview")
async def autofix_preview(file: Annotated[UploadFile, File(...)],
                          docx_rules: Annotated[str | None, Form()] = None):
    """Возвращает JSON с оригиналом и исправленным кодом для diff-просмотра."""
    content = await file.read()
    ext = _validate(file, content)
    name = file.filename or f"file{ext}"
    _require_autofix_support(ext)

    if ext == ".py":
        try:
            source = content.decode("utf-8")
        except UnicodeDecodeError:
            raise HTTPException(400, "Файл должен быть в UTF-8")
        result = autofix_python_code(source)
        return JSONResponse(content={
            "file_type": "python",
            "filename": name,
            "original": source,
            "fixed": result["fixed_code"],
            "changed": result["changed"],
        })

    # .docx — возвращаем сводку (полный diff невозможен для бинарного формата)
    custom_rules = _parse_rules(docx_rules)

    tmp = Path(tempfile.gettempdir()) / f"sd_prev_{uuid.uuid4().hex}.docx"
    try:
        tmp.write_bytes(content)
        # Проверяем до исправления
        report_before = check_docx_document(tmp, name, custom_rules)
        # Исправляем
        fixed_bytes = autofix_docx(tmp, custom_rules)
        # Сохраняем исправленный и проверяем снова
        tmp_fixed = Path(tempfile.gettempdir()) / f"sd_prevf_{uuid.uuid4().hex}.docx"
        try:
            tmp_fixed.write_bytes(fixed_bytes)
            report_after = check_docx_document(tmp_fixed, name, custom_rules)
        finally:
            try:
                os.remove(tmp_fixed)
            except OSError:
                pass

        return JSONResponse(content={
            "file_type": "docx",
            "filename": name,
            "before_issues": report_before["total_issues"],
            "after_issues": report_after["total_issues"],
            "before_summary": report_before["summary"],
            "after_summary": report_after["summary"],
            "changed": report_before["total_issues"] != report_after["total_issues"],
        })
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


@app.post("/api/autofix")
async def autofix(file: Annotated[UploadFile, File(...)],
                  docx_rules: Annotated[str | None, Form()] = None):
    content = await file.read()
    ext = _validate(file, content)
    name = file.filename or f"file{ext}"
    _require_autofix_support(ext)

    custom_rules = _parse_rules(docx_rules)

    try:
        if ext == ".py":
            try:
                source = content.decode("utf-8")
            except UnicodeDecodeError:
                raise HTTPException(400, "Файл должен быть в UTF-8")
            result = autofix_python_code(source)
            new_name = name.rsplit(".", 1)[0] + "_fixed.py"
            return Response(
                content=result["fixed_code"].encode("utf-8"),
                media_type="text/x-python; charset=utf-8",
                headers={"Content-Disposition": _content_disposition(new_name)},
            )

        # .docx
        tmp = Path(tempfile.gettempdir()) / f"sd_fix_{uuid.uuid4().hex}.docx"
        try:
            tmp.write_bytes(content)
            fixed = autofix_docx(tmp, custom_rules)
            new_name = name.rsplit(".", 1)[0] + "_fixed.docx"
            return Response(
                content=fixed,
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                headers={"Content-Disposition": _content_disposition(new_name)},
            )
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"Ошибка автоисправления: {e}")


@app.post("/api/export-pdf")
async def export_pdf(report: dict):
    required = {"file_type", "total_issues", "issues", "filename"}
    if not required.issubset(report.keys()):
        raise HTTPException(400, "Некорректный формат отчёта")
    try:
        pdf_bytes = generate_pdf_report(report)
    except Exception as e:
        raise HTTPException(500, f"Ошибка генерации PDF: {e}")
    filename = report["filename"].rsplit(".", 1)[0] + "_report.pdf"
    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": _content_disposition(filename)},
    )


@app.post("/api/generate-template")
async def api_generate_template(docx_rules: Annotated[str | None, Form()] = None):
    """Генерирует шаблон .docx по текущим правилам."""
    custom_rules = _parse_rules(docx_rules)
    try:
        template_bytes = generate_template(custom_rules)
    except Exception as e:
        raise HTTPException(500, f"Ошибка генерации шаблона: {e}")
    return Response(
        content=template_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": _content_disposition("Шаблон_АГУ.docx")},
    )


@app.get("/api/presets")
async def get_presets():
    """Возвращает доступные пресеты правил для .docx."""
    return JSONResponse(content=PRESETS)


@app.get("/api/languages")
async def get_languages():
    """Какие форматы принимаются и где доступно автоисправление."""
    return JSONResponse(content={
        "formats": [
            {
                "extension": ext,
                "name": LANGUAGE_NAMES.get(ext, "Документ Word"),
                "kind": "code" if ext in SUPPORTED_CODE_EXTENSIONS else "document",
                "autofix": ext in AUTOFIX_EXTENSIONS,
            }
            for ext in (
                [e for e in _EXT_ORDER if e in ALLOWED_EXTENSIONS]
                + sorted(ALLOWED_EXTENSIONS - set(_EXT_ORDER))
            )
        ],
        "max_file_size": MAX_FILE_SIZE,
    })


@app.get("/health")
async def health():
    return {"status": "ok", "service": "SmartDoc & Code Review", "version": "8.0.0"}
