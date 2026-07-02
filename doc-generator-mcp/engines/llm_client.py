"""LLM 客户端 —— 统一的模型调用接口"""

import json
import logging
import time
import re
from typing import Any

logger = logging.getLogger(__name__)


class LLMClient:
    """封装 LLM 调用，支持 OpenAI 兼容 API"""

    def __init__(self, config: dict):
        llm_cfg = config.get("llm", {})
        self.api_key = llm_cfg.get("api_key", "")
        self.base_url = llm_cfg.get("base_url", "")
        self.model = llm_cfg.get("model", "gpt-4o")
        self.temperature = llm_cfg.get("temperature", 0.3)
        self.max_retries = llm_cfg.get("max_retries", 3)
        self._client = None  # 缓存客户端实例

    def _get_client(self):
        """获取 OpenAI 客户端（缓存复用）"""
        if self._client is not None:
            return self._client

        try:
            from openai import OpenAI
        except ImportError:
            raise RuntimeError("openai 库未安装，请运行: pip install openai")

        kwargs = {"api_key": self.api_key}
        if self.base_url:
            kwargs["base_url"] = self.base_url
        self._client = OpenAI(**kwargs)
        return self._client

    def chat(self, messages: list[dict], response_format: dict | None = None) -> str:
        """发送聊天请求，返回文本响应（带重试）

        Args:
            messages: 消息列表
            response_format: 响应格式约束（如 JSON mode）

        Returns:
            模型响应文本
        """
        last_error = None
        for attempt in range(1, self.max_retries + 1):
            try:
                client = self._get_client()
                kwargs = {
                    "model": self.model,
                    "messages": messages,
                    "temperature": self.temperature,
                }
                if response_format:
                    kwargs["response_format"] = response_format

                response = client.chat.completions.create(**kwargs)
                return response.choices[0].message.content

            except Exception as e:
                last_error = e
                if attempt < self.max_retries:
                    wait = min(2 ** attempt, 10)  # 指数退避：2s, 4s, 8s, max 10s
                    logger.warning(
                        "LLM 调用失败 (attempt %d/%d): %s，%ds 后重试",
                        attempt, self.max_retries, e, wait,
                    )
                    time.sleep(wait)
                else:
                    logger.error("LLM 调用全部失败 (%d 次): %s", self.max_retries, e)

        raise RuntimeError(f"LLM 调用失败（已重试 {self.max_retries} 次）: {last_error}")

    def chat_json(self, messages: list[dict]) -> dict:
        """发送聊天请求，期望返回 JSON

        Args:
            messages: 消息列表

        Returns:
            解析后的 JSON 字典
        """
        raw = self.chat(messages, response_format={"type": "json_object"})

        # 尝试解析 JSON
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            # 尝试从 markdown 代码块中提取
            match = re.search(r'```(?:json)?\s*\n(.*?)\n```', raw, re.DOTALL)
            if match:
                try:
                    return json.loads(match.group(1))
                except json.JSONDecodeError:
                    pass
            raise ValueError(f"LLM 返回的内容不是有效 JSON: {raw[:200]}...")

    def structured_extract(self, text: str, target_schema: dict) -> dict:
        """从文本中提取结构化数据

        Args:
            text: 源文本
            target_schema: 目标 JSON Schema

        Returns:
            提取的结构化数据
        """
        messages = [
            {
                "role": "system",
                "content": (
                    "你是一个结构化数据提取专家。根据用户提供的文本和 JSON Schema，"
                    "提取并返回符合 Schema 的 JSON 数据。只返回 JSON，不要其他内容。"
                ),
            },
            {
                "role": "user",
                "content": f"文本内容:\n{text}\n\n目标 JSON Schema:\n{json.dumps(target_schema, ensure_ascii=False, indent=2)}",
            },
        ]
        return self.chat_json(messages)

    @property
    def is_available(self) -> bool:
        """检查 LLM 是否可用（有 API key）"""
        return bool(self.api_key)
