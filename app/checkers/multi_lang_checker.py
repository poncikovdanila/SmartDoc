"""Модуль проверки кода на нескольких языках.

Поддерживаемые языки и инструменты:
    .py   → flake8 (PEP 8 + pyflakes)              — модуль code_checker
    .js   → встроенный чекер на основе регулярок
    .sql  → встроенный чекер SQL-стиля
    .java → встроенный чекер Java-конвенций
    .cpp / .c / .h → встроенный чекер C/C++ стиля

Почему встроенные, а не внешние линтеры (ESLint, checkstyle и т. п.)?
    Для учебного проекта важна простота установки: pip install и готово.
    ESLint требует Node.js, checkstyle — JDK. Это лишний барьер. Наши
    встроенные чекеры покрывают 80% типичных студенческих ошибок: именование,
    пробелы, длина строк, скобки, комментарии. Для серьёзного продакшна
    можно подключить настоящие линтеры позже.

Ключевой приём против ложных срабатываний — маскирование (`_mask_line`):
перед тем как применять регулярки, из строки вычищаются строковые литералы
и комментарии. Длина строки при этом сохраняется, поэтому номера колонок
остаются верными. Без этого `var` в комментарии или `==` внутри строки
попадали бы в отчёт как настоящие замечания.

Каждый чекер возвращает единый формат, совместимый с code_checker:
filename, file_type, language_name, total_issues, summary, verdict, issues,
source_lines.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Callable, Iterable


# ═══════ Общие утилиты ═══════

def _read_source(file_path: Path) -> list[str]:
    try:
        return file_path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return []


def _mask_line(
    line: str,
    in_block: bool,
    *,
    line_comments: Iterable[str] = ("//",),
    quotes: str = "\"'",
    backslash_escape: bool = True,
) -> tuple[str, bool]:
    """Заменяет строковые литералы и комментарии пробелами.

    Длина результата всегда равна длине исходной строки — значит, позиция
    совпадения в замаскированной строке совпадает с колонкой в исходной.

    Args:
        line: исходная строка кода.
        in_block: находимся ли мы внутри блочного комментария /* ... */.
        line_comments: маркеры однострочного комментария ("//" или "--").
        quotes: символы, открывающие строковый литерал.
        backslash_escape: экранирование обратным слэшем (C-подобные языки).
            Если False — используется SQL-правило удвоенной кавычки ('').

    Returns:
        Пара (замаскированная строка, флаг «остались внутри блока»).
    """
    out: list[str] = []
    i = 0
    n = len(line)
    markers = tuple(line_comments)

    while i < n:
        if in_block:
            if line.startswith("*/", i):
                in_block = False
                out.append("  ")
                i += 2
            else:
                out.append(" ")
                i += 1
            continue

        if line.startswith("/*", i):
            in_block = True
            out.append("  ")
            i += 2
            continue

        if any(line.startswith(marker, i) for marker in markers):
            out.append(" " * (n - i))
            break

        ch = line[i]
        if ch in quotes:
            quote = ch
            out.append(" ")
            i += 1
            while i < n:
                if backslash_escape and line[i] == "\\" and i + 1 < n:
                    out.append("  ")
                    i += 2
                    continue
                if line[i] == quote:
                    # SQL: удвоенная кавычка внутри литерала — это экранирование
                    if not backslash_escape and line.startswith(quote * 2, i):
                        out.append("  ")
                        i += 2
                        continue
                    out.append(" ")
                    i += 1
                    break
                out.append(" ")
                i += 1
            continue

        out.append(ch)
        i += 1

    return "".join(out), in_block


def _add(
    issues: list[dict[str, Any]],
    line: int,
    column: int,
    code: str,
    message: str,
    description: str,
    severity: str,
) -> None:
    issues.append({
        "line": line,
        "column": column,
        "code": code,
        "message": message,
        "description": description,
        "severity": severity,
    })


def _check_line_length(
    issues: list[dict[str, Any]],
    line_no: int,
    line: str,
    limit: int,
    code: str,
) -> None:
    if len(line) > limit:
        _add(
            issues, line_no, limit + 1, code,
            f"Строка длиннее {limit} символов ({len(line)})",
            f"Рекомендуемый лимит строки — {limit} символов",
            "low",
        )


def _check_trailing_space(
    issues: list[dict[str, Any]],
    line_no: int,
    line: str,
    code: str,
) -> None:
    if line and line.rstrip() != line:
        _add(
            issues, line_no, len(line.rstrip()) + 1, code,
            "Пробелы в конце строки",
            "Удалите лишние пробелы в конце строки",
            "low",
        )


def _make_report(
    lines: list[str],
    issues: list[dict[str, Any]],
    filename: str,
    language: str,
) -> dict[str, Any]:
    """Собирает отчёт в формате, совместимом с code_checker.check_python_code."""
    issues.sort(key=lambda item: (item["line"], item.get("column", 0)))

    # Дедупликация: один и тот же код на одной строке показываем один раз.
    seen: set[tuple[int, str]] = set()
    deduped: list[dict[str, Any]] = []
    for issue in issues:
        key = (issue["line"], issue["code"])
        if key not in seen:
            seen.add(key)
            deduped.append(issue)

    summary = {"high": 0, "medium": 0, "low": 0}
    for issue in deduped:
        severity = issue.get("severity", "medium")
        if severity not in summary:
            severity = "medium"
            issue["severity"] = severity
        summary[severity] += 1

    # Вердикт по той же шкале, что и для Python: одно критичное замечание
    # работу не «валит», два и больше — уже повод доработать.
    if not deduped:
        verdict = "good"
    elif summary["high"] <= 1:
        verdict = "ok"
    else:
        verdict = "bad"

    return {
        "filename": filename,
        "file_type": language,
        "language_name": LANGUAGE_DISPLAY_NAMES.get(language, language),
        "total_issues": len(deduped),
        "summary": summary,
        "verdict": verdict,
        "issues": deduped,
        "source_lines": lines,
    }


# ═══════ JavaScript (.js) ═══════

# (код, серьёзность, регулярка, сообщение, описание)
JS_RULES: list[tuple[str, str, re.Pattern, str, str]] = [
    ("JS001", "medium", re.compile(r"^\t"),
     "Отступ табуляцией",
     "Используйте пробелы для отступов, не табуляцию"),
    ("JS002", "medium", re.compile(r"(?<![=!<>])={2}(?!=)"),
     "Используется == вместо ===",
     "Строгое сравнение === не приводит типы и потому безопаснее"),
    ("JS003", "medium", re.compile(r"!=(?!=)"),
     "Используется != вместо !==",
     "Строгое неравенство !== не приводит типы и потому безопаснее"),
    ("JS005", "high", re.compile(r"\bvar\b"),
     "Используется var",
     "Используйте let или const — у них блочная область видимости, "
     "поэтому переменная не «утекает» из блока"),
    ("JS006", "medium", re.compile(r"\bconsole\.(?:log|warn|error|debug|info)\s*\("),
     "Остался вызов console",
     "Уберите отладочный вывод в консоль перед сдачей работы"),
    ("JS008", "medium", re.compile(r"\bfunction\s+[A-Z]\w*\s*\("),
     "Функция начинается с заглавной буквы",
     "Функции именуются в camelCase. Заглавная буква — для классов "
     "и функций-конструкторов"),
    ("JS009", "high", re.compile(r"\beval\s*\("),
     "Используется eval",
     "eval выполняет произвольный код и потому небезопасен. "
     "Почти всегда есть замена без него"),
    ("JS011", "high", re.compile(r"\bdebugger\b"),
     "Остался оператор debugger",
     "debugger останавливает выполнение в браузере. Уберите его перед сдачей"),
    ("JS012", "medium", re.compile(r"==\s*(?:null|undefined)\b"),
     "Сравнение с null/undefined через ==",
     "Используйте === null, либо явную проверку на оба значения"),
]

JS_LINE_LENGTH = 100


def check_javascript(file_path: Path, original_filename: str) -> dict[str, Any]:
    """Проверяет .js-файл по типовым конвенциям современного JavaScript."""
    lines = _read_source(file_path)
    issues: list[dict[str, Any]] = []
    in_block = False

    for line_no, line in enumerate(lines, 1):
        masked, in_block = _mask_line(line, in_block, quotes="\"'`")

        for code, severity, pattern, message, description in JS_RULES:
            match = pattern.search(masked)
            if match:
                _add(issues, line_no, match.start() + 1, code,
                     message, description, severity)

        _check_trailing_space(issues, line_no, line, "JS004")
        _check_line_length(issues, line_no, line, JS_LINE_LENGTH, "JS010")

    return _make_report(lines, issues, original_filename, "javascript")


# ═══════ SQL (.sql) ═══════

SQL_KEYWORDS = {
    "select", "from", "where", "join", "inner", "outer", "left", "right",
    "on", "group", "by", "order", "having", "insert", "into", "values",
    "update", "set", "delete", "create", "table", "alter", "drop", "index",
    "and", "or", "not", "in", "between", "like", "is", "null", "as",
    "distinct", "union", "all", "exists", "case", "when", "then", "else", "end",
    "limit", "offset", "asc", "desc", "count", "sum", "avg", "min", "max",
}

SQL_LINE_LENGTH = 120

SQL_WORD_RE = re.compile(r"\b[A-Za-z_]+\b")
SQL_SELECT_STAR_RE = re.compile(r"\bSELECT\s+\*", re.IGNORECASE)
SQL_DELETE_NO_WHERE_RE = re.compile(r"^\s*DELETE\s+FROM\s+\w+\s*;?\s*$", re.IGNORECASE)
SQL_UPDATE_NO_WHERE_RE = re.compile(r"^\s*UPDATE\s+\w+\s+SET\b(?!.*\bWHERE\b)", re.IGNORECASE)


def check_sql(file_path: Path, original_filename: str) -> dict[str, Any]:
    """Проверяет .sql-файл: регистр ключевых слов, SELECT *, опасные запросы."""
    lines = _read_source(file_path)
    issues: list[dict[str, Any]] = []
    in_block = False

    for line_no, line in enumerate(lines, 1):
        masked, in_block = _mask_line(
            line, in_block, line_comments=("--", "#"),
            quotes="'\"", backslash_escape=False,
        )

        _check_trailing_space(issues, line_no, line, "SQL003")
        _check_line_length(issues, line_no, line, SQL_LINE_LENGTH, "SQL010")

        if not masked.strip():
            continue

        # Ключевые слова принято писать заглавными. Одно замечание на строку,
        # иначе на длинном запросе отчёт раздувается.
        for match in SQL_WORD_RE.finditer(masked):
            word = match.group()
            if word.lower() in SQL_KEYWORDS and word != word.upper():
                _add(issues, line_no, match.start() + 1, "SQL001",
                     f"Ключевое слово «{word}» не в верхнем регистре",
                     "SQL-ключевые слова принято писать ЗАГЛАВНЫМИ: "
                     "SELECT, FROM, WHERE",
                     "medium")
                break

        match = SQL_SELECT_STAR_RE.search(masked)
        if match:
            _add(issues, line_no, match.start() + 1, "SQL002",
                 "Используется SELECT *",
                 "Перечисляйте нужные столбцы явно — SELECT * тянет лишние "
                 "данные и ломается при изменении схемы таблицы",
                 "medium")

        if SQL_DELETE_NO_WHERE_RE.match(masked):
            _add(issues, line_no, 1, "SQL004",
                 "DELETE без условия WHERE",
                 "Такой запрос удалит все строки таблицы. "
                 "Добавьте WHERE, если это не задумано",
                 "high")

        if SQL_UPDATE_NO_WHERE_RE.match(masked):
            _add(issues, line_no, 1, "SQL005",
                 "UPDATE без условия WHERE",
                 "Такой запрос изменит все строки таблицы. "
                 "Добавьте WHERE, если это не задумано",
                 "high")

    return _make_report(lines, issues, original_filename, "sql")


# ═══════ Java (.java) ═══════

JAVA_LINE_LENGTH = 120

JAVA_CLASS_RE = re.compile(
    r"^\s*(?:(?:public|private|protected|final|abstract|static)\s+)*"
    r"class\s+([a-z]\w*)"
)
JAVA_METHOD_RE = re.compile(
    r"^\s*(?:(?:public|private|protected|static|final|synchronized|abstract)\s+)+"
    r"[\w<>\[\],.\s]+?\s+([A-Z]\w*)\s*\("
)
JAVA_EMPTY_CATCH_RE = re.compile(r"\bcatch\s*\([^)]*\)\s*\{\s*\}")


def check_java(file_path: Path, original_filename: str) -> dict[str, Any]:
    """Проверяет .java-файл по конвенциям именования и оформления Java."""
    lines = _read_source(file_path)
    issues: list[dict[str, Any]] = []
    in_block = False

    for line_no, line in enumerate(lines, 1):
        masked, in_block = _mask_line(line, in_block)

        _check_trailing_space(issues, line_no, line, "JAVA006")
        _check_line_length(issues, line_no, line, JAVA_LINE_LENGTH, "JAVA010")

        if line.startswith("\t"):
            _add(issues, line_no, 1, "JAVA004",
                 "Отступ табуляцией",
                 "Используйте 4 пробела для отступов",
                 "medium")

        stripped = masked.strip()
        if not stripped:
            continue

        if stripped == "{":
            _add(issues, line_no, masked.index("{") + 1, "JAVA001",
                 "Открывающая скобка на отдельной строке",
                 "В Java принят стиль K&R: открывающая скобка остаётся "
                 "на строке оператора",
                 "low")

        match = JAVA_CLASS_RE.match(masked)
        if match:
            _add(issues, line_no, match.start(1) + 1, "JAVA002",
                 f"Имя класса «{match.group(1)}» начинается со строчной буквы",
                 "Имена классов в Java пишутся в PascalCase (с заглавной буквы)",
                 "high")

        match = JAVA_METHOD_RE.match(masked)
        if match and not JAVA_CLASS_RE.match(masked):
            _add(issues, line_no, match.start(1) + 1, "JAVA003",
                 f"Имя метода «{match.group(1)}» начинается с заглавной буквы",
                 "Методы в Java именуются в camelCase (со строчной буквы)",
                 "medium")

        match = re.search(r"\bSystem\.(?:out|err)\.print", masked)
        if match:
            _add(issues, line_no, match.start() + 1, "JAVA005",
                 "Остался отладочный вывод System.out.println",
                 "Уберите отладочный вывод. Для журналирования используйте Logger",
                 "medium")

        match = JAVA_EMPTY_CATCH_RE.search(masked)
        if match:
            _add(issues, line_no, match.start() + 1, "JAVA007",
                 "Пустой блок catch",
                 "Проглоченное исключение прячет ошибку. Обработайте его "
                 "или хотя бы запишите в журнал",
                 "high")

    return _make_report(lines, issues, original_filename, "java")


# ═══════ C/C++ (.c, .cpp, .h) ═══════

CPP_LINE_LENGTH = 100

CPP_STD_HEADERS = (
    "iostream", "string", "vector", "map", "set", "algorithm",
    "cmath", "cstdlib", "cstdio", "cstring", "fstream", "sstream",
)
CPP_STD_INCLUDE_RE = re.compile(
    r'^\s*#\s*include\s+"(' + "|".join(CPP_STD_HEADERS) + r')"'
)
CPP_USING_STD_RE = re.compile(r"^\s*using\s+namespace\s+std\s*;")
CPP_PRINTF_RE = re.compile(r"\b(?:printf|scanf)\s*\(")
CPP_GOTO_RE = re.compile(r"\bgoto\b")
CPP_MALLOC_RE = re.compile(r"\b(?:malloc|calloc|realloc|free)\s*\(")


def check_cpp(file_path: Path, original_filename: str) -> dict[str, Any]:
    """Проверяет .c/.cpp/.h-файл по типовым конвенциям C++."""
    lines = _read_source(file_path)
    issues: list[dict[str, Any]] = []
    in_block = False
    is_cpp = Path(original_filename).suffix.lower() != ".c"

    for line_no, line in enumerate(lines, 1):
        masked, in_block = _mask_line(line, in_block)

        _check_trailing_space(issues, line_no, line, "CPP006")
        _check_line_length(issues, line_no, line, CPP_LINE_LENGTH, "CPP010")

        if line.startswith("\t"):
            _add(issues, line_no, 1, "CPP004",
                 "Отступ табуляцией",
                 "Используйте пробелы для отступов",
                 "medium")

        # Директива #include проверяется по исходной строке: маскирование
        # вычистило бы имя заголовка вместе с кавычками.
        match = CPP_STD_INCLUDE_RE.match(line)
        if match:
            _add(issues, line_no, match.start(1) + 1, "CPP002",
                 f'#include "{match.group(1)}" — нужны угловые скобки',
                 "Стандартные заголовки подключаются через <>, а не кавычки: "
                 f"#include <{match.group(1)}>",
                 "medium")

        if not masked.strip():
            continue

        if CPP_USING_STD_RE.match(masked):
            _add(issues, line_no, 1, "CPP001",
                 "using namespace std",
                 "using namespace std загрязняет глобальное пространство имён "
                 "и провоцирует конфликты. Пишите std::cout, std::string",
                 "high")

        if is_cpp:
            match = CPP_PRINTF_RE.search(masked)
            if match:
                _add(issues, line_no, match.start() + 1, "CPP003",
                     "Используется printf/scanf",
                     "В C++ предпочтительнее std::cout / std::cin — "
                     "они типобезопасны",
                     "low")

            match = CPP_MALLOC_RE.search(masked)
            if match:
                _add(issues, line_no, match.start() + 1, "CPP007",
                     "Используется malloc/free",
                     "В C++ управление памятью делают через new/delete, "
                     "а лучше через умные указатели",
                     "medium")

        match = CPP_GOTO_RE.search(masked)
        if match:
            _add(issues, line_no, match.start() + 1, "CPP005",
                 "Используется goto",
                 "goto усложняет чтение кода. Используйте циклы и функции",
                 "high")

    return _make_report(lines, issues, original_filename, "cpp")


# ═══════ Роутер: определяет язык по расширению ═══════

CHECKERS: dict[str, Callable[[Path, str], dict[str, Any]]] = {
    ".js": check_javascript,
    ".sql": check_sql,
    ".java": check_java,
    ".cpp": check_cpp,
    ".c": check_cpp,
    ".h": check_cpp,
}

# .py проверяется модулем code_checker, поэтому в CHECKERS его нет.
SUPPORTED_CODE_EXTENSIONS = {".py"} | set(CHECKERS)

LANGUAGE_NAMES = {
    ".py": "Python", ".js": "JavaScript", ".sql": "SQL",
    ".java": "Java", ".cpp": "C++", ".c": "C", ".h": "C/C++ Header",
}

# file_type из отчёта → человекочитаемое имя языка
LANGUAGE_DISPLAY_NAMES = {
    "python": "Python", "javascript": "JavaScript", "sql": "SQL",
    "java": "Java", "cpp": "C/C++",
}

# Автоисправление пока есть только для Python (autopep8) и .docx
AUTOFIX_EXTENSIONS = {".py", ".docx"}


def check_code_file(file_path: Path, original_filename: str, extension: str) -> dict[str, Any]:
    """Роутер: вызывает нужный чекер по расширению файла."""
    checker = CHECKERS.get(extension)
    if checker is None:
        raise ValueError(f"Нет чекера для расширения {extension}")
    return checker(file_path, original_filename)
