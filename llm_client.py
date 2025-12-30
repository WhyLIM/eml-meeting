import json
import time
from typing import Dict, Optional, List
import requests
from config import API_CONFIGS, MAX_RETRIES, REQUEST_TIMEOUT


class LLMClient:

    def __init__(self, api_name: str = "zai-plan"):
        if api_name not in API_CONFIGS:
            raise ValueError(
                f"不支持的API: {api_name}. 支持的API: {list(API_CONFIGS.keys())}")

        self.config = API_CONFIGS[api_name]
        self.api_name = api_name
        self.api_type = self.config["type"]

        if not self.config["api_key"]:
            raise ValueError(f"未设置API密钥: {api_name}_API_KEY")

    def _create_anthropic_request(self, messages: List[Dict[str, str]],
                                  **kwargs) -> Dict:
        system_prompt = kwargs.get("system", "")
        messages_list = []

        if system_prompt:
            messages_list.append({"role": "user", "content": system_prompt})

        for msg in messages:
            if msg["role"] != "system":
                messages_list.append(msg)

        return {
            "model": kwargs.get("model", self.config["model"]),
            "max_tokens": kwargs.get("max_tokens", 4096),
            "messages": messages_list
        }

    def _create_openai_request(self, messages: List[Dict[str, str]],
                               **kwargs) -> Dict:
        return {
            "model": kwargs.get("model", self.config["model"]),
            "messages": messages,
            "max_tokens": kwargs.get("max_tokens", 4096),
            "temperature": kwargs.get("temperature", 0.0)
        }

    def _create_zai_request(self, messages: List[Dict[str, str]],
                            **kwargs) -> Dict:
        return self._create_openai_request(messages, **kwargs)

    def _create_gemini_request(self, messages: List[Dict[str, str]],
                               **kwargs) -> Dict:
        contents = []

        system_prompt = kwargs.get("system", "")
        if system_prompt:
            contents.append({
                "role": "user",
                "parts": [{
                    "text": system_prompt
                }]
            })

        for msg in messages:
            if msg["role"] != "system":
                contents.append({
                    "role": msg["role"],
                    "parts": [{
                        "text": msg["content"]
                    }]
                })

        return {
            "contents": contents,
            "generationConfig": {
                "maxOutputTokens": kwargs.get("max_tokens", 4096),
                "temperature": kwargs.get("temperature", 0.0)
            }
        }

    def _call_api(self, request_data: Dict) -> Dict:
        headers = {}

        if self.api_type == "anthropic":
            headers = {
                "Content-Type": "application/json",
                "x-api-key": self.config["api_key"],
                "anthropic-version": "2023-06-01",
                "anthropic-dangerous-direct-browser-access": "true"
            }
        elif self.api_type == "openai":
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.config['api_key']}"
            }
        elif self.api_type == "zai":
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.config['api_key']}"
            }
        elif self.api_type == "gemini":
            headers = {"Content-Type": "application/json"}

        url = self.config["url"]

        if self.api_type == "gemini":
            url = f"{url}/models/{self.config['model']}:generateContent?key={self.config['api_key']}"
        elif self.api_type == "openai":
            url = f"{url}/chat/completions"
        elif self.api_type == "zai":
            url = f"{url}/chat/completions"
        elif self.api_type == "anthropic":
            url = f"{url}/v1/messages"

        response = requests.post(url,
                                 headers=headers,
                                 json=request_data,
                                 timeout=REQUEST_TIMEOUT)

        response.raise_for_status()
        return response.json()

    def _parse_response(self, response: Dict) -> str:
        try:
            if self.api_type == "anthropic":
                return response["content"][0]["text"]
            elif self.api_type == "openai":
                return response["choices"][0]["message"]["content"]
            elif self.api_type == "zai":
                return response["choices"][0]["message"]["content"]
            elif self.api_type == "gemini":
                return response["candidates"][0]["content"]["parts"][0]["text"]
        except (KeyError, IndexError) as e:
            raise ValueError(f"解析API响应失败: {e}, 响应: {response}")

    def chat(self, messages: List[Dict[str, str]], **kwargs) -> str:
        request_data = None

        if self.api_type == "anthropic":
            request_data = self._create_anthropic_request(messages, **kwargs)
        elif self.api_type == "openai":
            request_data = self._create_openai_request(messages, **kwargs)
        elif self.api_type == "zai":
            request_data = self._create_zai_request(messages, **kwargs)
        elif self.api_type == "gemini":
            request_data = self._create_gemini_request(messages, **kwargs)
        else:
            raise ValueError(f"不支持的API类型: {self.api_type}")

        last_error = None
        for attempt in range(MAX_RETRIES):
            try:
                response = self._call_api(request_data)
                return self._parse_response(response)
            except requests.exceptions.RequestException as e:
                last_error = e
                if attempt < MAX_RETRIES - 1:
                    wait_time = 2**attempt
                    time.sleep(wait_time)
                else:
                    raise
            except Exception as e:
                last_error = e
                if attempt < MAX_RETRIES - 1:
                    wait_time = 2**attempt
                    time.sleep(wait_time)
                else:
                    raise

        raise last_error if last_error else Exception("API调用失败")

    def extract_json(self, text: str) -> Optional[Dict]:
        import re

        json_pattern = r'\[[^\[\]]*(?:\[[^\[\]]*\][^\[\]]*)*\]'
        matches = re.findall(json_pattern, text, re.DOTALL)

        for match in matches:
            try:
                return json.loads(match)
            except json.JSONDecodeError:
                continue

        return None


def get_available_apis() -> List[str]:
    return list(API_CONFIGS.keys())
