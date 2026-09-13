"""Модуль проверки Python-кода на соответствие стандарту PEP 8.

Использует flake8 (комбайн pycodestyle + pyflakes + mccabe). Запускается через
subprocess, чтобы избежать конфликтов глобального состояния flake8 и обеспечить
изоляцию между запросами.
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Any

# Описания самых частых кодов flake8/PEP 8 на русском, чтобы отчёт был понятнее
# студенту. Не претендует на полноту — только то, что встречается чаще всего.
RULE_DESCRIPTIONS: dict[str, str] = {
    # pycodestyle (E/W) — собственно PEP 8
    "E101": "Несовместимое смешение табуляции и пробелов в отступе",
    "E111": "Отступ не кратен четырём пробелам",
    "E114": "Отступ не кратен четырём (комментарий)",
    "E117": "Слишком большой отступ",
    "E121": "Отступ продолжения строки не выровнен",
    "E122": "Отсутствует отступ продолжения строки",
    "E125": "Отступ продолжения строки не отличается от следующего блока",
    "E127": "Отступ продолжения строки переусерднствует с выравниванием",
    "E128": "Отступ продолжения строки выровнен не по открывающему разделителю",
    "E201": "Лишний пробел после открывающей скобки",
    "E202": "Лишний пробел перед закрывающей скобкой",
    "E203": "Лишний пробел перед двоеточием/запятой/точкой с запятой",
    "E211": "Лишний пробел перед скобкой",
    "E225": "Отсутствует пробел вокруг оператора",
    "E226": "Отсутствует пробел вокруг арифметического оператора",
    "E227": "Отсутствует пробел вокруг побитового оператора",
    "E228": "Отсутствует пробел вокруг оператора по модулю",
    "E231": "Отсутствует пробел после запятой/двоеточия",
    "E251": "Лишний пробел вокруг = в параметрах функции",
    "E261": "Перед строчным комментарием должно быть минимум 2 пробела",
    "E262": "Строчный комментарий должен начинаться с '# '",
    "E265": "Блочный комментарий должен начинаться с '# '",
    "E266": "Слишком много решёток для блочного комментария",
    "E301": "Ожидалась 1 пустая строка",
    "E302": "Ожидалось 2 пустые строки между функциями/классами верхнего уровня",
    "E303": "Слишком много пустых строк подряд",
    "E305": "Ожидалось 2 пустые строки после определения класса/функции",
    "E306": "Ожидалась 1 пустая строка перед вложенным определением",
    "E401": "Несколько импортов в одной строке",
    "E402": "Импорт модуля не в начале файла",
    "E501": "Слишком длинная строка (> 79 символов)",
    "E502": "Обратный слэш — лишний внутри скобок",
    "E701": "Несколько инструкций в одной строке (двоеточие)",
    "E702": "Несколько инструкций в одной строке (точка с запятой)",
    "E703": "Лишняя точка с запятой в конце инструкции",
    "E711": "Сравнение с None должно быть через 'is' / 'is not'",
    "E712": "Сравнение с True/False должно быть через 'is' / 'is not'",
    "E713": "Проверку отсутствия вхождения пишут как 'not in'",
    "E714": "Проверку 'не является' пишут как 'is not'",
    "E721": "Тип сравнивайте через isinstance(), а не ==",
    "E722": "Не используйте голый except:",
    "E741": "Неоднозначное имя переменной (l, I, O)",
    "W191": "Отступ табуляцией (PEP 8 требует пробелы)",
    "W291": "Конечные пробелы в строке",
    "W292": "Нет символа новой строки в конце файла",
    "W293": "Пустая строка содержит пробелы",
    "W391": "Лишние пустые строки в конце файла",
    "W503": "Перенос строки перед бинарным оператором",
    "W504": "Перенос строки после бинарного оператора",
    "W605": "Некорректная escape-последовательность",
    # pyflakes (F)
    "F401": "Импортируется, но не используется",
    "F403": "Импорт через 'from X import *' скрывает имена",
    "F405": "Имя может быть определено из импорта 'from X import *'",
    "F811": "Переопределение неиспользуемой переменной",
    "F821": "Использование необъявленного имени",
    "F841": "Локальная переменная присвоена, но не используется",
}

# Серьёзность: E5xx (длина строки) и W — низкая, F (логика) — высокая, остальное — средняя.
SEVERITY_HIGH = {"F401", "F403", "F405", "F811", "F821", "F841", "E722", "E741"}
SEVERITY_LOW_PREFIXES = ("W",)


def _severity_for(code: str) -> str:
    if code in SEVERITY_HIGH:
        return "high"
    if any(code.startswith(prefix) for prefix in SEVERITY_LOW_PREFIXES):
        return "low"
    return "medium"


# Регулярка под формат вывода flake8: "путь:строка:колонка: КОД сообщение".
FLAKE8_LINE_RE = re.compile(
    r"^(?P<path>.+?):(?P<line>\d+):(?P<col>\d+):\s*(?P<code>[A-Z]\d+)\s+(?P<msg>.*)$"
)


def check_python_code(file_path: Path, original_filename: str) -> dict[str, Any]:
    """Проверяет .py-файл через flake8 и возвращает структурированный отчёт.

    Args:
        file_path: путь к временной копии загруженного файла на диске.
        original_filename: исходное имя файла, чтобы показать его в отчёте.

    Returns:
        Словарь с полями: filename, file_type, total_issues, summary, issues, source_lines.
    """
    # Запускаем flake8 как отдельный процесс. Параметры:
    #   --max-line-length=79 — каноническое значение из PEP 8.
    #   --no-show-source — мы сами покажем код в отчёте.
    # Возвращаемый код != 0 при наличии замечаний — это нормально, не падаем.
    try:
        result = subprocess.run(
            [
                sys.executable,
                "-m",
                "flake8",
                "--max-line-length=79",
                "--no-show-source",
                str(file_path),
            ],
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return {
            "filename": original_filename,
            "file_type": "python",
            "total_issues": 0,
            "summary": {"high": 0, "medium": 0, "low": 0},
            "issues": [],
            "error": "Анализ кода занял слишком много времени и был прерван (>30 с).",
            "source_lines": [],
        }

    issues: list[dict[str, Any]] = []
    summary = {"high": 0, "medium": 0, "low": 0}

    for raw_line in result.stdout.splitlines():
        match = FLAKE8_LINE_RE.match(raw_line.strip())
        if not match:
            continue
        code = match.group("code")
        severity = _severity_for(code)
        summary[severity] += 1
        issues.append(
            {
                "line": int(match.group("line")),
                "column": int(match.group("col")),
                "code": code,
                "message": match.group("msg"),
                "description": RULE_DESCRIPTIONS.get(code, match.group("msg")),
                "severity": severity,
            }
        )

    # Сортируем по строке, потом по колонке
    issues.sort(key=lambda item: (item["line"], item["column"]))

    # Дедупликация: если на одной строке один и тот же код — оставляем только первый.
    # flake8 иногда выдаёт дубли, например для E225 на строке с несколькими операторами.
    seen: set[tuple[int, str]] = set()
    deduped: list[dict[str, Any]] = []
    for issue in issues:
        key = (issue["line"], issue["code"])
        if key not in seen:
            seen.add(key)
            deduped.append(issue)
    issues = deduped

    # Пересчитываем summary после дедупликации
    summary = {"high": 0, "medium": 0, "low": 0}
    for issue in issues:
        summary[issue["severity"]] += 1

    # Подгружаем исходник для отображения «контекста» возле каждой ошибки.
    try:
        source_lines = file_path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        source_lines = []

    # Вердикт
    # good — замечаний нет; ok — без критичных или одно критичное (одно поле
    # в разделе приложений не должно «валить» работу целиком); bad — иначе
    if not issues:
        verdict = "good"
    elif summary["high"] <= 1:
        verdict = "ok"
    else:
        verdict = "bad"

    return {
        "filename": original_filename,
        "file_type": "python",
        "total_issues": len(issues),
        "summary": summary,
        "verdict": verdict,
        "issues": issues,
        "source_lines": source_lines,
    }


def check_python_code_as_json(file_path: Path, original_filename: str) -> str:
    """Удобная обёртка для CLI-использования."""
    return json.dumps(
        check_python_code(file_path, original_filename), ensure_ascii=False, indent=2
    )
