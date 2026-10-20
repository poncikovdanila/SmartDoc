"""Сменный слой провайдеров языковой модели.

Зачем он нужен. Сервисом пользуются разные люди с разными возможностями:
у кого-то есть ключ Claude, кому-то важнее цена и он возьмёт DeepSeek или
Qwen, а кто-то не хочет отправлять документы наружу вообще и поднимет
модель у себя через Ollama. Привязка к одному поставщику закрывала бы
две трети этих случаев.

Поэтому код работает с одним интерфейсом, а конкретный поставщик
задаётся переменными окружения:

    AI_PROVIDER=anthropic           # по умолчанию, если задан ANTHROPIC_API_KEY
    ANTHROPIC_API_KEY=sk-ant-...
    AI_MODEL=claude-opus-5-5        # необязательно

    AI_PROVIDER=openai              # любой OpenAI-совместимый API
    AI_BASE_URL=https://api.deepseek.com
    AI_API_KEY=...
    AI_MODEL=deepseek-chat

    AI_PROVIDER=openai              # локальная модель, ключ не нужен
    AI_BASE_URL=http://localhost:11434/v1
    AI_MODEL=qwen2.5:14b

Два протокола отличаются формой вызова инструментов и структурированного
ответа, поэтому у каждого свой адаптер. Наружу оба отдают одинаковый
результат — класс Reply.
"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from typing import Any

# Значения по умолчанию для каждого протокола. Переопределяются AI_MODEL.
DEFAULT_MODELS = {
    "anthropic": "claude-opus-5-5",
    "openai": "gpt-4o-mini",
}

TIMEOUT_SECONDS = 90.0
MAX_TOKENS = 8000


class AINotConfigured(RuntimeError):
    """Слой ИИ не настроен или не хватает зависимостей."""


@dataclass
class ToolCall:
    """Запрос модели вызвать инструмент."""

    id: str
    name: str
    arguments: dict[str, Any]


@dataclass
class Reply:
    """Единый ответ модели независимо от поставщика."""

    text: str = ""
    tool_calls: list[ToolCall] = field(default_factory=list)
    # Ответ модели в формате поставщика — его нужно вернуть обратно
    # на следующем шаге диалога, иначе связь с вызовом инструмента теряется
    raw_message: Any = None


def _env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def active_provider() -> str | None:
    """Какой поставщик настроен. None, если слой ИИ выключен."""
    explicit = _env("AI_PROVIDER").lower()
    if explicit in ("anthropic", "openai"):
        # Локальной модели ключ не нужен, поэтому для openai достаточно адреса
        if explicit == "openai" and (_env("AI_API_KEY") or _env("AI_BASE_URL")):
            return "openai"
        if explicit == "anthropic" and _env("ANTHROPIC_API_KEY"):
            return "anthropic"
        return None
    if _env("ANTHROPIC_API_KEY"):
        return "anthropic"
    if _env("AI_API_KEY") or _env("AI_BASE_URL"):
        return "openai"
    return None


def is_configured() -> bool:
    return active_provider() is not None


def model_name() -> str:
    provider = active_provider()
    if provider is None:
        return ""
    return _env("AI_MODEL") or DEFAULT_MODELS[provider]


class AnthropicProvider:
    """Поставщик поверх официального SDK Anthropic."""

    name = "anthropic"

    def __init__(self) -> None:
        try:
            import anthropic
        except ImportError as exc:
            raise AINotConfigured(
                "Выбран провайдер anthropic, но пакет не установлен. "
                "Выполните: pip install anthropic"
            ) from exc
        self._client = anthropic.Anthropic(timeout=TIMEOUT_SECONDS, max_retries=1)
        self._model = model_name()

    def complete(
        self,
        system: str,
        messages: list[dict[str, Any]],
        tools: list[dict[str, Any]] | None = None,
    ) -> Reply:
        kwargs: dict[str, Any] = {
            "model": self._model,
            "max_tokens": MAX_TOKENS,
            "system": system,
            "messages": messages,
        }
        if tools:
            kwargs["tools"] = [
                {
                    "name": tool["name"],
                    "description": tool["description"],
                    "input_schema": tool["parameters"],
                }
                for tool in tools
            ]

        response = self._client.messages.create(**kwargs)

        text_parts: list[str] = []
        calls: list[ToolCall] = []
        for block in response.content:
            if block.type == "text":
                text_parts.append(block.text)
            elif block.type == "tool_use":
                calls.append(ToolCall(id=block.id, name=block.name, arguments=dict(block.input)))

        return Reply(
            text="\n".join(text_parts).strip(),
            tool_calls=calls,
            raw_message={"role": "assistant", "content": response.content},
        )

    @staticmethod
    def tool_result_messages(results: list[tuple[str, str]]) -> list[dict[str, Any]]:
        """В этом протоколе все результаты идут одним сообщением пользователя."""
        return [{
            "role": "user",
            "content": [
                {"type": "tool_result", "tool_use_id": call_id, "content": output}
                for call_id, output in results
            ],
        }]

    def parse(self, system: str, prompt: str, schema: dict[str, Any]) -> dict[str, Any]:
        """Структурированный ответ по схеме JSON."""
        response = self._client.messages.create(
            model=self._model,
            max_tokens=MAX_TOKENS,
            system=system,
            messages=[{"role": "user", "content": prompt}],
            output_config={"format": {"type": "json_schema", "schema": schema}},
        )
        text = next(b.text for b in response.content if b.type == "text")
        return json.loads(text)


class OpenAICompatibleProvider:
    """Поставщик для любого API, говорящего на протоколе OpenAI.

    Сюда попадают DeepSeek, Qwen через DashScope, OpenRouter и локальные
    модели под Ollama или vLLM. Протокол один, отличаются только адрес
    и название модели.
    """

    name = "openai"

    def __init__(self) -> None:
        try:
            from openai import OpenAI
        except ImportError as exc:
            raise AINotConfigured(
                "Выбран OpenAI-совместимый провайдер, но пакет не установлен. "
                "Выполните: pip install openai"
            ) from exc
        self._client = OpenAI(
            # Локальной модели ключ не нужен, но клиент требует непустую строку
            api_key=_env("AI_API_KEY") or "not-needed",
            base_url=_env("AI_BASE_URL") or None,
            timeout=TIMEOUT_SECONDS,
            max_retries=1,
        )
        self._model = model_name()

    def complete(
        self,
        system: str,
        messages: list[dict[str, Any]],
        tools: list[dict[str, Any]] | None = None,
    ) -> Reply:
        kwargs: dict[str, Any] = {
            "model": self._model,
            "max_tokens": MAX_TOKENS,
            "messages": [{"role": "system", "content": system}] + messages,
        }
        if tools:
            kwargs["tools"] = [
                {
                    "type": "function",
                    "function": {
                        "name": tool["name"],
                        "description": tool["description"],
                        "parameters": tool["parameters"],
                    },
                }
                for tool in tools
            ]

        response = self._client.chat.completions.create(**kwargs)
        message = response.choices[0].message

        calls: list[ToolCall] = []
        for call in (message.tool_calls or []):
            try:
                arguments = json.loads(call.function.arguments or "{}")
            except (ValueError, TypeError):
                # Модель вернула невалидный JSON в аргументах. Это не повод
                # ронять диалог: инструмент получит пустые аргументы и сам
                # сообщит, что ему не хватает данных.
                arguments = {}
            calls.append(ToolCall(id=call.id, name=call.function.name, arguments=arguments))

        return Reply(
            text=(message.content or "").strip(),
            tool_calls=calls,
            raw_message=message.model_dump(exclude_none=True),
        )

    @staticmethod
    def tool_result_messages(results: list[tuple[str, str]]) -> list[dict[str, Any]]:
        """В этом протоколе каждый результат идёт отдельным сообщением."""
        return [
            {"role": "tool", "tool_call_id": call_id, "content": output}
            for call_id, output in results
        ]

    def parse(self, system: str, prompt: str, schema: dict[str, Any]) -> dict[str, Any]:
        response = self._client.chat.completions.create(
            model=self._model,
            max_tokens=MAX_TOKENS,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": prompt},
            ],
            response_format={
                "type": "json_schema",
                "json_schema": {"name": "result", "schema": schema, "strict": True},
            },
        )
        return json.loads(response.choices[0].message.content)


def get_provider() -> AnthropicProvider | OpenAICompatibleProvider:
    """Создаёт поставщика по настройкам окружения.

    Raises:
        AINotConfigured: если слой ИИ выключен или не хватает зависимостей.
    """
    provider = active_provider()
    if provider is None:
        raise AINotConfigured(
            "Слой ИИ не настроен. Задайте ANTHROPIC_API_KEY либо "
            "AI_PROVIDER=openai вместе с AI_BASE_URL и AI_MODEL."
        )
    return AnthropicProvider() if provider == "anthropic" else OpenAICompatibleProvider()
