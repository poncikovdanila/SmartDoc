"""Тесты предпросмотра автоисправления (/api/autofix-preview).

Эндпоинт отвечает на вопрос «что случится с файлом, если нажать
„Исправить“» — до того, как файл изменён и скачан. Для кода это полный
текст до и после, для .docx — количество замечаний до и после, потому что
построчный diff для архива с разметкой смысла не имеет.
"""
from __future__ import annotations

from pathlib import Path

import pytest

pytest.importorskip("httpx", reason="TestClient требует httpx")

from fastapi.testclient import TestClient  # noqa: E402

from app.main import app  # noqa: E402

client = TestClient(app)

DEMO_DOCX = Path(__file__).resolve().parent.parent / "sample_files" / "demo_document_full.docx"

# Голый except и импорты в одну строку — autopep8 с autoflake это чинят
FIXABLE_PY = b"""import os, sys


def f():
    try:
        return 1
    except:
        return None
"""

# Неоднозначное имя переменной (E741) — замечание есть, но автоматически
# его не исправить: выбор нового имени за автором
UNFIXABLE_PY = b"def f(l):\n    return l\n"


def _preview(name: str, content: bytes):
    return client.post("/api/autofix-preview",
                       files={"file": (name, content, "text/plain")})


class TestPreviewPython:
    def test_возвращает_оригинал_и_результат(self):
        body = _preview("a.py", FIXABLE_PY).json()
        assert body["file_type"] == "python"
        assert body["filename"] == "a.py"
        assert body["original"] == FIXABLE_PY.decode("utf-8")
        assert body["changed"] is True
        assert body["fixed"] != body["original"]

    def test_исправление_действительно_снимает_замечания(self):
        before = client.post("/api/check", files={"file": ("a.py", FIXABLE_PY, "text/plain")}).json()
        fixed = _preview("a.py", FIXABLE_PY).json()["fixed"]
        after = client.post(
            "/api/check",
            files={"file": ("a.py", fixed.encode("utf-8"), "text/plain")}).json()
        assert after["total_issues"] < before["total_issues"]

    def test_нечего_исправлять(self):
        body = _preview("a.py", UNFIXABLE_PY).json()
        assert body["changed"] is False
        assert body["fixed"] == body["original"]

    def test_предпросмотр_не_меняет_исходный_файл(self):
        # Два вызова подряд должны дать один и тот же ответ: эндпоинт
        # ничего не сохраняет и не накапливает состояние между запросами
        first = _preview("a.py", FIXABLE_PY).json()
        second = _preview("a.py", FIXABLE_PY).json()
        assert first == second

    def test_не_utf8(self):
        response = _preview("a.py", b"x = '\xff\xfe'\n")
        assert response.status_code == 400


class TestPreviewDocx:
    @pytest.mark.skipif(not DEMO_DOCX.exists(), reason="нет демо-документа")
    def test_сводка_до_и_после(self):
        response = client.post(
            "/api/autofix-preview",
            files={"file": (DEMO_DOCX.name, DEMO_DOCX.read_bytes(),
                            "application/vnd.openxmlformats-officedocument."
                            "wordprocessingml.document")})
        assert response.status_code == 200
        body = response.json()
        assert body["file_type"] == "docx"
        for field in ("before_issues", "after_issues",
                      "before_summary", "after_summary", "changed"):
            assert field in body, f"нет поля {field}"
        # Исправление не должно добавлять замечаний
        assert body["after_issues"] <= body["before_issues"]
        for level in ("high", "medium", "low"):
            assert body["after_summary"][level] <= body["before_summary"][level], level

    @pytest.mark.skipif(not DEMO_DOCX.exists(), reason="нет демо-документа")
    def test_сводка_согласована_с_общим_числом(self):
        body = client.post(
            "/api/autofix-preview",
            files={"file": (DEMO_DOCX.name, DEMO_DOCX.read_bytes(),
                            "application/vnd.openxmlformats-officedocument."
                            "wordprocessingml.document")}).json()
        assert sum(body["before_summary"].values()) == body["before_issues"]
        assert sum(body["after_summary"].values()) == body["after_issues"]


class TestPreviewLimits:
    @pytest.mark.parametrize("name,content", [
        ("a.js", b"var x = 1;\n"),
        ("a.sql", b"select * from t;\n"),
        ("a.java", b"public class report {}\n"),
        ("a.cpp", b"using namespace std;\n"),
    ])
    def test_языки_без_автоисправления(self, name, content):
        response = _preview(name, content)
        assert response.status_code == 422
        assert "не поддерживается" in response.json()["detail"]

    def test_неизвестный_формат(self):
        assert _preview("a.rb", b"puts 1\n").status_code == 400

    def test_пустой_файл(self):
        assert _preview("a.py", b"").status_code == 400
