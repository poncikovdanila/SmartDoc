"""Подготовка JSON-схемы для структурированного ответа.

Pydantic описывает вложенные модели через $defs и не ставит
additionalProperties. Строгий режим у OpenAI-совместимых поставщиков
требует, чтобы у каждого объекта было additionalProperties: false и
перечислены все обязательные поля. Эта функция приводит схему к виду,
который принимают оба протокола.
"""
from __future__ import annotations

from typing import Any


def strict_schema(model: Any) -> dict[str, Any]:
    """JSON-схема модели Pydantic, пригодная для строгого режима."""
    return _tighten(model.model_json_schema())


def _tighten(node: Any) -> Any:
    if isinstance(node, list):
        return [_tighten(item) for item in node]
    if not isinstance(node, dict):
        return node

    result = {key: _tighten(value) for key, value in node.items()}
    if result.get("type") == "object" and "properties" in result:
        result.setdefault("additionalProperties", False)
        # Строгий режим требует перечислить все поля как обязательные:
        # необязательных в нём не бывает
        result["required"] = list(result["properties"].keys())
    return result
