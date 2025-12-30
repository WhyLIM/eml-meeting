import os
from dotenv import load_dotenv

load_dotenv()

ZAI_PLAN_API_URL = "https://open.bigmodel.cn/api/coding/paas/v4"
ZAI_API_URL = "https://open.bigmodel.cn/api/paas/v4/"
OPENAI_API_URL = "https://api.openai.com/v1"
DEEPSEEK_API_URL = "https://api.deepseek.com/v1"
GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta"

ZAI_API_KEY = os.getenv("ZAI_API_KEY", "")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

API_CONFIGS = {
    "zai-plan": {
        "url": ZAI_PLAN_API_URL,
        "api_key": ZAI_API_KEY,
        "model": "glm-4.5",
        "type": "zai",
        "max_concurrency": 5
    },
    "zai": {
        "url": ZAI_API_URL,
        "api_key": ZAI_API_KEY,
        "model": "glm-4.5",
        "type": "zai",
        "max_concurrency": 5
    },
    "openai": {
        "url": OPENAI_API_URL,
        "api_key": OPENAI_API_KEY,
        "model": "gpt-4o",
        "type": "openai",
        "max_concurrency": 5
    },
    "deepseek": {
        "url": DEEPSEEK_API_URL,
        "api_key": DEEPSEEK_API_KEY,
        "model": "deepseek-chat",
        "type": "openai",
        "max_concurrency": 5
    },
    "gemini": {
        "url": GEMINI_API_URL,
        "api_key": GEMINI_API_KEY,
        "model": "gemini-3-flash",
        "type": "gemini",
        "max_concurrency": 5
    }
}

DEFAULT_API = "zai-plan"
DEFAULT_MODEL = API_CONFIGS[DEFAULT_API]["model"]

MAX_RETRIES = 3
REQUEST_TIMEOUT = 120

DEFAULT_MAX_CONCURRENCY = 3


def get_available_apis():
    return list(API_CONFIGS.keys())


def get_max_concurrency(api_name: str) -> int:
    if api_name in API_CONFIGS:
        return API_CONFIGS[api_name].get("max_concurrency",
                                         DEFAULT_MAX_CONCURRENCY)
    return DEFAULT_MAX_CONCURRENCY
