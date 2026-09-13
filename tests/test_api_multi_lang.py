"""Интеграционные тесты HTTP-слоя: маршрутизация файлов по расширению.

Проверяем не сами правила (для этого есть test_multi_lang.py), а то, что
запрос доходит до нужного чекера и что API честно отвечает там, где
возможности нет — например, при попытке автоисправить JavaScript.
"""
from __future__ import annotations

import pytest

pytest.importorskip("httpx", reason="TestClient требует httpx")

from fastapi.testclient import TestClient  # noqa: E402

from app.main import app  # noqa: E402

client = TestClient(app)


def _check(name: str, source: str):
    return client.post("/api/check", files={"file": (name, source.encode("utf-8"),
                                                     "text/plain")})


class TestLanguagesEndpoint:
    def test_перечисляет_все_форматы(self):
        data = client.get("/api/languages").json()
        extensions = {f["extension"] for f in data["formats"]}
        assert {".docx", ".py", ".js", ".sql", ".java", ".cpp", ".c", ".h"} == extensions

    def test_отмечает_где_есть_автоисправление(self):
        formats = {f["extension"]: f for f in client.get("/api/languages").json()["formats"]}
        assert formats[".py"]["autofix"] is True
        assert formats[".docx"]["autofix"] is True
        assert formats[".js"]["autofix"] is False

    def test_различает_код_и_документ(self):
        formats = {f["extension"]: f for f in client.get("/api/languages").json()["formats"]}
        assert formats[".docx"]["kind"] == "document"
        assert formats[".java"]["kind"] == "code"

    def test_версия_в_health(self):
        assert client.get("/health").json()["version"] == "8.0.0"


class TestCheckRouting:
    @pytest.mark.parametrize("name,source,file_type", [
        ("a.py", "import os\n", "python"),
        ("a.js", "var x = 1;\n", "javascript"),
        ("a.sql", "select * from t;\n", "sql"),
        ("a.java", "public class report {}\n", "java"),
        ("a.cpp", "using namespace std;\n", "cpp"),
        ("a.c", "goto done;\n", "cpp"),
        ("a.h", "using namespace std;\n", "cpp"),
    ])
    def test_файл_уходит_в_свой_чекер(self, name, source, file_type):
        response = _check(name, source)
        assert response.status_code == 200
        body = response.json()
        assert body["file_type"] == file_type
        assert body["total_issues"] >= 1
        assert body["verdict"] in {"good", "ok", "bad"}

    def test_расширение_в_верхнем_регистре(self):
        assert _check("A.JS", "var x = 1;\n").json()["file_type"] == "javascript"

    def test_неподдерживаемый_формат(self):
        response = _check("a.rb", "puts 1\n")
        assert response.status_code == 400
        assert ".js" in response.json()["detail"]

    def test_пустой_файл_отклоняется(self):
        response = client.post("/api/check",
                               files={"file": ("a.js", b"", "text/plain")})
        assert response.status_code == 400

    def test_чистый_код_без_замечаний(self):
        source = "const total = 1;\nexport default total;\n"
        body = _check("clean.js", source).json()
        assert body["total_issues"] == 0
        assert body["verdict"] == "good"


class TestAutofixLimits:
    @pytest.mark.parametrize("name,source", [
        ("a.js", "var x = 1;\n"),
        ("a.sql", "select * from t;\n"),
        ("a.java", "public class report {}\n"),
        ("a.cpp", "using namespace std;\n"),
    ])
    def test_автоисправление_отвечает_понятной_ошибкой(self, name, source):
        response = client.post(
            "/api/autofix",
            files={"file": (name, source.encode("utf-8"), "text/plain")})
        assert response.status_code == 422
        assert "не поддерживается" in response.json()["detail"]

    def test_python_по_прежнему_исправляется(self):
        response = client.post(
            "/api/autofix",
            files={"file": ("a.py", b"import os\nx=1\n", "text/plain")})
        assert response.status_code == 200
        assert b"x = 1" in response.content

    def test_предпросмотр_тоже_отклоняет(self):
        response = client.post(
            "/api/autofix-preview",
            files={"file": ("a.js", b"var x = 1;\n", "text/plain")})
        assert response.status_code == 422


class TestBatch:
    def test_разные_языки_в_одной_пачке(self):
        response = client.post("/api/check-batch", files=[
            ("files", ("a.js", b"var x = 1;\n", "text/plain")),
            ("files", ("b.java", b"public class report {}\n", "text/plain")),
            ("files", ("c.py", b"import os\n", "text/plain")),
            ("files", ("d.sql", b"select * from t;\n", "text/plain")),
        ])
        assert response.status_code == 200
        body = response.json()
        assert body["file_count"] == 4
        types = [r["file_type"] for r in body["reports"]]
        assert types == ["javascript", "java", "python", "sql"]
        assert body["total_issues"] == sum(r["total_issues"] for r in body["reports"])

    def test_плохой_файл_не_ломает_пачку(self):
        response = client.post("/api/check-batch", files=[
            ("files", ("a.js", b"var x = 1;\n", "text/plain")),
            ("files", ("b.rb", b"puts 1\n", "text/plain")),
        ])
        body = response.json()
        assert body["file_count"] == 2
        assert body["reports"][0]["file_type"] == "javascript"
        assert "error" in body["reports"][1]


class TestPdfExport:
    @pytest.mark.parametrize("name,source", [
        ("a.js", "var x = 1;\n"),
        ("a.sql", "select * from t;\n"),
        ("a.java", "public class report {}\n"),
    ])
    def test_отчёт_по_коду_выгружается_в_pdf(self, name, source):
        report = _check(name, source).json()
        response = client.post("/api/export-pdf", json=report)
        assert response.status_code == 200
        assert response.content.startswith(b"%PDF")
