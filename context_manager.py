"""对话历史管理：滑动窗口 + 简单 token 预算控制。"""
from __future__ import annotations
from typing import Any
from config import MAX_HISTORY_TURNS, MAX_CONTEXT_TOKENS


def _rough_tokens(text: str) -> int:
    """粗略估算 token 数：中文字符约 1 token，英文按空格分词。"""
    return max(len(text) // 2, 1)


def _message_tokens(msg: dict[str, Any]) -> int:
    content = msg.get("content") or ""
    if isinstance(content, list):
        content = " ".join(
            part.get("text", "") for part in content if isinstance(part, dict)
        )
    return _rough_tokens(str(content))


class ContextManager:
    def __init__(self) -> None:
        self._history: list[dict[str, Any]] = []

    def add(self, message: dict[str, Any]) -> None:
        self._history.append(message)
        self._trim()

    def get_history(self) -> list[dict[str, Any]]:
        return list(self._history)

    def _trim(self) -> None:
        # 保留最近 MAX_HISTORY_TURNS 轮（每轮含 user + assistant）
        max_messages = MAX_HISTORY_TURNS * 2
        if len(self._history) > max_messages:
            self._history = self._history[-max_messages:]

        # token 预算裁剪：从最旧消息开始丢弃
        while self._history and self._total_tokens() > MAX_CONTEXT_TOKENS:
            self._history.pop(0)

    def _total_tokens(self) -> int:
        return sum(_message_tokens(m) for m in self._history)

    def clear(self) -> None:
        self._history.clear()
