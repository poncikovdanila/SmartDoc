"""Ассистент: диалог поверх уже готового отчёта.

Чем он отличается от одиночных вызовов в bibliography и explain. Там
запрос уходит один раз и ответ приходит сразу. Здесь модель может
попросить вызвать функцию, получить результат и попросить следующую, пока
не соберёт достаточно сведений для ответа.

Два решения, которые стоит проговорить.

Состояния на сервере нет. Отчёт и вся переписка приходят от браузера
каждым запросом. Сервер ничего не хранит между вызовами: это и проще,
и честнее по отношению к чужим дипломам.

Инструменты только читают. Ассистент может посмотреть замечания, текст
абзаца, сводку и список литературы, но не может изменить документ. Пока
не станет ясно, что он отвечает толково, давать ему право править файл
рано. Добавить такой инструмент потом — несколько строк.
"""
from __future__ import annotations

import json
from typing import Any, Callable

from app.ai.providers import Reply, get_provider

# Сколько раз подряд модель может запросить инструменты, прежде чем мы
# потребуем ответа. Ограничение и от зацикливания, и от расходов.
MAX_TOOL_ROUNDS = 4

# Ограничения на размер выдачи инструментов: без них в модель уедет
# весь документ, что дорого и бессмысленно.
MAX_ISSUES_RETURNED = 30
MAX_PARAGRAPHS_RETURNED = 10
MAX_HISTORY_MESSAGES = 30

SYSTEM = """Ты помощник внутри SmartDoc — сервиса проверки оформления \
документов и кода.

Пользователь уже проверил файл и видит отчёт. Твоя задача — отвечать на \
вопросы по этому отчёту и помогать разобраться, что и почему исправлять.

Как работать:
- сведения бери инструментами, а не из общих соображений. Не помнишь \
точное значение — посмотри инструментом;
- не выдумывай замечаний, которых нет в отчёте, и не меняй их \
формулировки на свои;
- требования к оформлению у каждого вуза свои. Ориентируйся на то, что \
показывают инструменты, а не на то, что принято «вообще». Не называй \
стандарт, если он не указан в данных;
- менять документ ты не можешь. Если пользователь просит исправить, \
объясни, что делает кнопка «Исправить и скачать», и что придётся \
поправить руками.

Как писать:
- по-русски, спокойно и коротко. Без вступлений и ободрений;
- если замечаний много, начинай с того, что важнее: сначала критичные, \
потом остальные;
- ссылайся на номера абзацев, чтобы пользователь нашёл место;
- не знаешь ответа — так и скажи."""


# ───────── Описания инструментов ─────────

TOOLS: list[dict[str, Any]] = [
    {
        "name": "get_summary",
        "description": (
            "Сводка по проверенному файлу: имя, тип, число замечаний по "
            "важности, общий вердикт. С этого стоит начинать, если вопрос "
            "общий."
        ),
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "list_issues",
        "description": (
            "Список замечаний из отчёта. Можно отфильтровать по коду "
            "замечания или по важности."
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "code": {"type": "string", "description": "Код замечания, например FONT_MISMATCH"},
                "severity": {
                    "type": "string",
                    "enum": ["high", "medium", "low"],
                    "description": "Важность: high — критично, medium — замечание, low — мелочь",
                },
            },
            "required": [],
        },
    },
    {
        "name": "get_paragraph",
        "description": (
            "Текст абзаца документа по его номеру и все замечания, "
            "относящиеся к этому абзацу. Номера берутся из замечаний."
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "index": {"type": "integer", "description": "Номер абзаца из отчёта"},
            },
            "required": ["index"],
        },
    },
    {
        "name": "find_text",
        "description": "Найти абзацы документа, содержащие заданный текст.",
        "parameters": {
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "Что искать"},
            },
            "required": ["query"],
        },
    },
    {
        "name": "list_bibliography",
        "description": "Записи списка литературы из документа, если он найден.",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
]


# ───────── Реализация инструментов поверх отчёта ─────────

def _paragraph_blocks(report: dict[str, Any]) -> list[dict[str, Any]]:
    preview = report.get("document_preview") or {}
    return [b for b in preview.get("blocks", []) if b.get("kind") == "p"]


def build_tools(report: dict[str, Any]) -> dict[str, Callable[[dict[str, Any]], str]]:
    """Собирает инструменты, замкнутые на конкретный отчёт."""
    issues = report.get("issues") or []
    blocks = _paragraph_blocks(report)
    by_index = {b["index"]: b for b in blocks}

    def get_summary(_: dict[str, Any]) -> str:
        summary = report.get("summary") or {}
        return json.dumps({
            "файл": report.get("filename"),
            "тип": report.get("file_type"),
            "всего_замечаний": report.get("total_issues", 0),
            "критичных": summary.get("high", 0),
            "замечаний": summary.get("medium", 0),
            "мелочей": summary.get("low", 0),
            "вердикт": report.get("verdict"),
            "абзацев_проверено": report.get("paragraphs_checked"),
        }, ensure_ascii=False)

    def list_issues(args: dict[str, Any]) -> str:
        selected = issues
        code = args.get("code")
        severity = args.get("severity")
        if code:
            selected = [i for i in selected if i.get("code") == code]
        if severity:
            selected = [i for i in selected if i.get("severity") == severity]
        if not selected:
            return "По такому фильтру замечаний нет."
        payload = [
            {
                "код": i.get("code"),
                "важность": i.get("severity"),
                "место": i.get("location"),
                "абзац": i.get("para_index"),
                "строка": i.get("line"),
                "суть": i.get("description") or i.get("message"),
                "требуется": i.get("expected"),
                "фактически": i.get("actual"),
            }
            for i in selected[:MAX_ISSUES_RETURNED]
        ]
        return json.dumps({
            "найдено": len(selected),
            "показано": len(payload),
            "замечания": payload,
        }, ensure_ascii=False)

    def get_paragraph(args: dict[str, Any]) -> str:
        index = args.get("index")
        if not isinstance(index, int) or index not in by_index:
            return f"Абзаца с номером {index} в документе нет."
        related = [
            {"код": i.get("code"), "суть": i.get("description") or i.get("message")}
            for i in issues if i.get("para_index") == index
        ]
        return json.dumps({
            "номер": index,
            "текст": by_index[index].get("text", ""),
            "выравнивание": by_index[index].get("align"),
            "уровень_заголовка": by_index[index].get("heading"),
            "замечания": related,
        }, ensure_ascii=False)

    def find_text(args: dict[str, Any]) -> str:
        query = str(args.get("query") or "").strip().lower()
        if not query:
            return "Не задано, что искать."
        found = [
            {"номер": b["index"], "текст": b.get("text", "")[:200]}
            for b in blocks if query in (b.get("text") or "").lower()
        ]
        if not found:
            return "Ничего не найдено."
        return json.dumps({
            "найдено": len(found),
            "абзацы": found[:MAX_PARAGRAPHS_RETURNED],
        }, ensure_ascii=False)

    def list_bibliography(_: dict[str, Any]) -> str:
        entries = report.get("bibliography_entries") or []
        if not entries:
            return "Список литературы в документе не найден."
        return json.dumps({
            "записей": len(entries),
            "источники": [
                {"номер": e.get("number"), "текст": e.get("text")}
                for e in entries[:MAX_ISSUES_RETURNED]
            ],
        }, ensure_ascii=False)

    return {
        "get_summary": get_summary,
        "list_issues": list_issues,
        "get_paragraph": get_paragraph,
        "find_text": find_text,
        "list_bibliography": list_bibliography,
    }


# ───────── Диалог ─────────

def run_chat(
    messages: list[dict[str, Any]],
    report: dict[str, Any],
) -> dict[str, Any]:
    """Проводит один ход диалога, вызывая инструменты по запросу модели.

    Args:
        messages: переписка от браузера, роли user и assistant.
        report: отчёт о проверке, на который опираются инструменты.

    Returns:
        Ответ ассистента и список вызванных инструментов.
    """
    provider = get_provider()
    tools = build_tools(report)
    history = list(messages[-MAX_HISTORY_MESSAGES:])
    used: list[str] = []

    for _ in range(MAX_TOOL_ROUNDS):
        reply: Reply = provider.complete(SYSTEM, history, tools=TOOLS)
        if not reply.tool_calls:
            return {"reply": reply.text, "tools_used": used}

        results: list[tuple[str, str]] = []
        for call in reply.tool_calls:
            used.append(call.name)
            handler = tools.get(call.name)
            if handler is None:
                results.append((call.id, f"Инструмента {call.name} не существует."))
                continue
            try:
                results.append((call.id, handler(call.arguments)))
            except Exception as exc:  # инструмент не должен ронять диалог
                results.append((call.id, f"Ошибка инструмента: {exc}"))

        history.append(reply.raw_message)
        history.extend(provider.tool_result_messages(results))

    # Круги кончились, а модель всё ещё просит инструменты. Просим ответить
    # тем, что уже собрано, иначе диалог не завершится.
    history.append({
        "role": "user",
        "content": "Ответь на вопрос по уже собранным сведениям, не вызывая инструменты.",
    })
    final = provider.complete(SYSTEM, history, tools=None)
    return {"reply": final.text, "tools_used": used}
