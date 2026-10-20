"""Совместимость: слой ИИ переехал в providers.py.

Модуль оставлен, чтобы существующий код и тесты продолжали работать
с прежними именами. Вся логика выбора поставщика живёт в providers.
"""
from __future__ import annotations

from app.ai.providers import (  # noqa: F401
    AINotConfigured,
    active_provider,
    get_provider,
    is_configured,
    model_name,
)

# Уровень усилий применим только к моделям Claude; у остальных
# поставщиков параметра нет, и слой провайдеров его не передаёт.
EFFORT = "low"
