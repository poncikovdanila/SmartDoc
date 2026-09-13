"""Тесты чекеров для JavaScript, SQL, Java и C/C++.

Главный акцент — не на том, что «грязный» файл ловится (это просто), а на
том, что чистый файл НЕ ловится. Ложное срабатывание в учебном инструменте
дороже пропуска: студент перестаёт доверять отчёту.
"""
from __future__ import annotations

import pytest

from app.checkers.multi_lang_checker import (
    AUTOFIX_EXTENSIONS,
    LANGUAGE_NAMES,
    SUPPORTED_CODE_EXTENSIONS,
    _mask_line,
    check_code_file,
    check_cpp,
    check_java,
    check_javascript,
    check_sql,
)


def _write(tmp_path, name: str, source: str):
    path = tmp_path / name
    path.write_text(source, encoding="utf-8")
    return path


def _codes(report) -> set[str]:
    return {issue["code"] for issue in report["issues"]}


# ═══════ Маскирование строк и комментариев ═══════

class TestMaskLine:
    def test_строковый_литерал_вычищается(self):
        masked, _ = _mask_line('const a = "var x == 1";', False)
        assert "var" not in masked
        assert "==" not in masked
        assert masked.startswith("const a = ")

    def test_длина_строки_сохраняется(self):
        line = 'if (a == "text") { // var\n'.rstrip("\n")
        masked, _ = _mask_line(line, False)
        assert len(masked) == len(line)

    def test_однострочный_комментарий_вычищается(self):
        masked, _ = _mask_line("let x = 1; // var y == 2", False)
        assert "var" not in masked
        assert "let x = 1;" in masked

    def test_блочный_комментарий_переносится_на_следующую_строку(self):
        first, in_block = _mask_line("/* начало var", False)
        assert in_block is True
        assert "var" not in first
        second, in_block = _mask_line("ещё var */ let x = 1;", in_block)
        assert in_block is False
        assert "var" not in second
        assert "let x = 1;" in second

    def test_экранированная_кавычка_не_закрывает_литерал(self):
        masked, _ = _mask_line(r'const s = "a\" var b"; let c = 1;', False)
        assert "var" not in masked
        assert "let c = 1;" in masked

    def test_sql_удвоенная_кавычка(self):
        masked, _ = _mask_line(
            "WHERE name = 'it''s select' AND id = 1",
            False, line_comments=("--",), quotes="'\"", backslash_escape=False,
        )
        assert "select" not in masked
        assert "AND id = 1" in masked


# ═══════ JavaScript ═══════

CLEAN_JS = """const total = items.length;
if (total === 0) {
    return null;
}
// здесь только комментарий: var, ==, eval(
const msg = "строка с var, == и eval(x) внутри";
export default total;
"""


class TestJavaScript:
    def test_чистый_файл_без_замечаний(self, tmp_path):
        report = check_javascript(_write(tmp_path, "clean.js", CLEAN_JS), "clean.js")
        assert report["total_issues"] == 0, report["issues"]
        assert report["verdict"] == "good"

    def test_var_и_eval_критичны(self, tmp_path):
        src = "var x = 1;\neval('2+2');\n"
        report = check_javascript(_write(tmp_path, "a.js", src), "a.js")
        assert {"JS005", "JS009"} <= _codes(report)
        assert report["summary"]["high"] >= 2
        assert report["verdict"] == "bad"

    def test_нестрогое_сравнение(self, tmp_path):
        report = check_javascript(_write(tmp_path, "a.js", "if (a == b) {}\n"), "a.js")
        assert "JS002" in _codes(report)

    def test_строгое_сравнение_не_ловится(self, tmp_path):
        report = check_javascript(
            _write(tmp_path, "a.js", "if (a === b && c !== d) {}\n"), "a.js")
        assert _codes(report) == set()

    def test_стрелочная_функция_не_ловится(self, tmp_path):
        report = check_javascript(
            _write(tmp_path, "a.js", "const f = (a) => a + 1;\n"), "a.js")
        assert _codes(report) == set()

    def test_console_и_debugger(self, tmp_path):
        src = "console.log(1);\ndebugger;\n"
        report = check_javascript(_write(tmp_path, "a.js", src), "a.js")
        assert {"JS006", "JS011"} <= _codes(report)

    def test_функция_с_заглавной(self, tmp_path):
        report = check_javascript(
            _write(tmp_path, "a.js", "function Render() {}\n"), "a.js")
        assert "JS008" in _codes(report)

    def test_шаблонная_строка_маскируется(self, tmp_path):
        report = check_javascript(
            _write(tmp_path, "a.js", "const s = `var x == 1`;\n"), "a.js")
        assert _codes(report) == set()

    def test_длинная_строка(self, tmp_path):
        src = "const s = '" + "a" * 120 + "';\n"
        report = check_javascript(_write(tmp_path, "a.js", src), "a.js")
        assert "JS010" in _codes(report)

    def test_колонка_указывает_на_место_ошибки(self, tmp_path):
        report = check_javascript(
            _write(tmp_path, "a.js", "let a = 1; var b = 2;\n"), "a.js")
        issue = next(i for i in report["issues"] if i["code"] == "JS005")
        assert issue["line"] == 1
        assert issue["column"] == 12


# ═══════ SQL ═══════

CLEAN_SQL = """-- комментарий: select from where
SELECT id, name
FROM users
WHERE name = 'select * from users' AND active = 1;
"""


class TestSQL:
    def test_чистый_файл_без_замечаний(self, tmp_path):
        report = check_sql(_write(tmp_path, "clean.sql", CLEAN_SQL), "clean.sql")
        assert report["total_issues"] == 0, report["issues"]
        assert report["verdict"] == "good"

    def test_строчные_ключевые_слова(self, tmp_path):
        report = check_sql(
            _write(tmp_path, "a.sql", "select id from users;\n"), "a.sql")
        assert "SQL001" in _codes(report)

    def test_смешанный_регистр_тоже_замечание(self, tmp_path):
        # Раньше эта ветка молча пропускалась из-за перевёрнутого условия
        report = check_sql(
            _write(tmp_path, "a.sql", "Select id From users;\n"), "a.sql")
        assert "SQL001" in _codes(report)
        issue = next(i for i in report["issues"] if i["code"] == "SQL001")
        assert "Select" in issue["message"]

    def test_select_звёздочка(self, tmp_path):
        report = check_sql(
            _write(tmp_path, "a.sql", "SELECT * FROM users;\n"), "a.sql")
        assert "SQL002" in _codes(report)

    def test_delete_без_where_критично(self, tmp_path):
        report = check_sql(_write(tmp_path, "a.sql", "DELETE FROM logs;\n"), "a.sql")
        issue = next(i for i in report["issues"] if i["code"] == "SQL004")
        assert issue["severity"] == "high"

    def test_delete_с_where_не_ловится(self, tmp_path):
        report = check_sql(
            _write(tmp_path, "a.sql", "DELETE FROM logs WHERE id = 1;\n"), "a.sql")
        assert "SQL004" not in _codes(report)

    def test_update_без_where_критично(self, tmp_path):
        report = check_sql(
            _write(tmp_path, "a.sql", "UPDATE users SET active = 0;\n"), "a.sql")
        assert "SQL005" in _codes(report)

    def test_update_с_where_не_ловится(self, tmp_path):
        report = check_sql(
            _write(tmp_path, "a.sql", "UPDATE users SET active = 0 WHERE id = 1;\n"),
            "a.sql")
        assert "SQL005" not in _codes(report)

    def test_одно_замечание_о_регистре_на_строку(self, tmp_path):
        report = check_sql(
            _write(tmp_path, "a.sql", "select id from users where id = 1;\n"), "a.sql")
        assert len([i for i in report["issues"] if i["code"] == "SQL001"]) == 1


# ═══════ Java ═══════

CLEAN_JAVA = """public class Report {
    private int count;

    /* блочный комментарий: class report, System.out.print */
    public int getCount() {
        return count;
    }

    public String describe() {
        return "class report and System.out.print inside a string";
    }
}
"""


class TestJava:
    def test_чистый_файл_без_замечаний(self, tmp_path):
        report = check_java(_write(tmp_path, "Report.java", CLEAN_JAVA), "Report.java")
        assert report["total_issues"] == 0, report["issues"]

    def test_класс_со_строчной(self, tmp_path):
        report = check_java(
            _write(tmp_path, "a.java", "public class report {}\n"), "a.java")
        issue = next(i for i in report["issues"] if i["code"] == "JAVA002")
        assert issue["severity"] == "high"

    def test_метод_с_заглавной(self, tmp_path):
        src = "public class A {\n    public void PrintAll() {}\n}\n"
        report = check_java(_write(tmp_path, "a.java", src), "a.java")
        assert "JAVA003" in _codes(report)

    def test_обычный_метод_не_ловится(self, tmp_path):
        src = "public class A {\n    public void printAll() {}\n}\n"
        report = check_java(_write(tmp_path, "a.java", src), "a.java")
        assert "JAVA003" not in _codes(report)

    def test_конструкция_new_не_считается_методом(self, tmp_path):
        src = "public class A {\n    private Helper h = new Helper();\n}\n"
        report = check_java(_write(tmp_path, "a.java", src), "a.java")
        assert "JAVA003" not in _codes(report)

    def test_отладочный_вывод(self, tmp_path):
        report = check_java(
            _write(tmp_path, "a.java", "System.out.println(1);\n"), "a.java")
        assert "JAVA005" in _codes(report)

    def test_пустой_catch(self, tmp_path):
        report = check_java(
            _write(tmp_path, "a.java", "try { f(); } catch (Exception e) {}\n"),
            "a.java")
        issue = next(i for i in report["issues"] if i["code"] == "JAVA007")
        assert issue["severity"] == "high"

    def test_скобка_на_отдельной_строке(self, tmp_path):
        src = "public class A\n{\n}\n"
        report = check_java(_write(tmp_path, "a.java", src), "a.java")
        assert "JAVA001" in _codes(report)

    def test_табуляция(self, tmp_path):
        report = check_java(
            _write(tmp_path, "a.java", "\tint x = 1;\n"), "a.java")
        assert "JAVA004" in _codes(report)


# ═══════ C/C++ ═══════

CLEAN_CPP = """#include <iostream>

// комментарий: goto, malloc, using namespace std
int main() {
    std::cout << "goto malloc printf внутри строки" << std::endl;
    return 0;
}
"""


class TestCpp:
    def test_чистый_файл_без_замечаний(self, tmp_path):
        report = check_cpp(_write(tmp_path, "main.cpp", CLEAN_CPP), "main.cpp")
        assert report["total_issues"] == 0, report["issues"]

    def test_using_namespace_std(self, tmp_path):
        report = check_cpp(
            _write(tmp_path, "a.cpp", "using namespace std;\n"), "a.cpp")
        issue = next(i for i in report["issues"] if i["code"] == "CPP001")
        assert issue["severity"] == "high"

    def test_стандартный_заголовок_в_кавычках(self, tmp_path):
        report = check_cpp(
            _write(tmp_path, "a.cpp", '#include "iostream"\n'), "a.cpp")
        assert "CPP002" in _codes(report)

    def test_свой_заголовок_в_кавычках_допустим(self, tmp_path):
        report = check_cpp(
            _write(tmp_path, "a.cpp", '#include "myheader.h"\n'), "a.cpp")
        assert "CPP002" not in _codes(report)

    def test_goto(self, tmp_path):
        report = check_cpp(_write(tmp_path, "a.cpp", "goto done;\n"), "a.cpp")
        assert "CPP005" in _codes(report)

    def test_printf_в_си_не_замечание(self, tmp_path):
        # В чистом C printf — норма, а не стилевая ошибка
        report = check_cpp(_write(tmp_path, "a.c", 'printf("x");\n'), "a.c")
        assert "CPP003" not in _codes(report)

    def test_printf_в_плюсах_замечание(self, tmp_path):
        report = check_cpp(_write(tmp_path, "a.cpp", 'printf("x");\n'), "a.cpp")
        assert "CPP003" in _codes(report)

    def test_malloc_в_плюсах(self, tmp_path):
        report = check_cpp(
            _write(tmp_path, "a.cpp", "int *p = (int*)malloc(4);\n"), "a.cpp")
        assert "CPP007" in _codes(report)

    def test_заголовочный_файл_проверяется_как_плюсы(self, tmp_path):
        report = check_cpp(
            _write(tmp_path, "a.h", "using namespace std;\n"), "a.h")
        assert "CPP001" in _codes(report)


# ═══════ Формат отчёта и роутер ═══════

class TestReportFormat:
    @pytest.mark.parametrize("name,source", [
        ("a.js", "var x = 1;\n"),
        ("a.sql", "select * from t;\n"),
        ("a.java", "public class report {}\n"),
        ("a.cpp", "using namespace std;\n"),
    ])
    def test_обязательные_поля(self, tmp_path, name, source):
        path = _write(tmp_path, name, source)
        report = check_code_file(path, name, path.suffix)
        for field in ("filename", "file_type", "language_name", "total_issues",
                      "summary", "verdict", "issues", "source_lines"):
            assert field in report, f"нет поля {field}"
        assert report["filename"] == name
        assert report["verdict"] in {"good", "ok", "bad"}
        assert report["source_lines"] == source.splitlines()

    @pytest.mark.parametrize("name,source", [
        ("a.js", "var x = 1;\nvar y = 2;\nvar z = 3;\n"),
        ("a.sql", "select * from t;\ndelete from t;\n"),
    ])
    def test_сводка_совпадает_с_числом_замечаний(self, tmp_path, name, source):
        path = _write(tmp_path, name, source)
        report = check_code_file(path, name, path.suffix)
        assert sum(report["summary"].values()) == report["total_issues"]
        assert report["total_issues"] == len(report["issues"])

    def test_замечания_отсортированы_по_строке(self, tmp_path):
        src = "var c = 3;\n" * 5
        path = _write(tmp_path, "a.js", src)
        report = check_code_file(path, "a.js", ".js")
        lines = [i["line"] for i in report["issues"]]
        assert lines == sorted(lines)

    def test_дубли_одного_кода_на_строке_схлопываются(self, tmp_path):
        path = _write(tmp_path, "a.js", "var a = 1; var b = 2;\n")
        report = check_code_file(path, "a.js", ".js")
        assert len([i for i in report["issues"] if i["code"] == "JS005"]) == 1

    def test_пустой_файл(self, tmp_path):
        path = _write(tmp_path, "a.js", "")
        report = check_code_file(path, "a.js", ".js")
        assert report["total_issues"] == 0
        assert report["verdict"] == "good"

    def test_неизвестное_расширение(self, tmp_path):
        path = _write(tmp_path, "a.rb", "puts 1\n")
        with pytest.raises(ValueError):
            check_code_file(path, "a.rb", ".rb")

    def test_список_расширений_согласован(self):
        assert ".py" in SUPPORTED_CODE_EXTENSIONS
        assert AUTOFIX_EXTENSIONS == {".py", ".docx"}
        for ext in SUPPORTED_CODE_EXTENSIONS:
            assert ext in LANGUAGE_NAMES, f"нет названия для {ext}"
