"""Подключение к Claude API.

Слой необязательный. Без переменной окружения ANTHROPIC_API_KEY весь
основной функционал работает как прежде, а эндпоинты с ИИ отвечают, что
он не настроен. Это сознательно: детерминированные проверки оформления
не должны зависеть от внешнего сервиса, который может быть недоступен,
стоить денег или отвечать по-разному на один и тот же документ.

Что модели НЕ поручается: поля, кегль, межстрочный интервал, красная
строка. Там python-docx даёт точное значение, и языковая модель внесла бы
только разброд.
"""
from __future__ import annotations

import os
from functools import lru_cache
from typing import Any

# Модель общего назначения: задачи здесь короткие, но требуют понимания
# правил оформления, а не простой классификации.
MODEL = "claude-opus-5-5"

# Низкий уровень усилий: разбор одной библиографической записи или
# объяснение замечания не требуют долгих рассуждений.
EFFORT = "low"

# Запрос должен укладываться в разумное время: пользователь ждёт ответа
# в браузере, а не в фоновой очереди.
TIMEOUT_SECONDS = 60.0


class AINotConfigured(RuntimeError):
    """Ключ API не задан — функции с ИИ недоступны."""


def is_configured() -> bool:
    """Задан ли ключ API."""
    return bool(os.environ.get("ANTHROPIC_API_KEY", "").strip())


@lru_cache(maxsize=1)
def _build_client() -> Any:
    try:
        import anthropic
    except ImportError as exc:  # pragma: no cover - зависит от окружения
        # Ключ задан, а пакета нет. Без явного сообщения это выглядело бы
        # как загадочная ошибка сервера.
        raise AINotConfigured(
            "Ключ ANTHROPIC_API_KEY задан, но пакет anthropic не установлен. "
            "Выполните: pip install anthropic"
        ) from exc
    return anthropic.Anthropic(timeout=TIMEOUT_SECONDS, max_retries=1)


def get_client() -> Any:
    """Возвращает клиента Claude API.

    Raises:
        AINotConfigured: если ключ не задан.
    """
    if not is_configured():
        raise AINotConfigured(
            "Проверки с участием ИИ не настроены: не задан ANTHROPIC_API_KEY"
        )
    return _build_client()
