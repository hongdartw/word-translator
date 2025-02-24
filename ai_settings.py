from enum import Enum
import os
from dotenv import load_dotenv

# 載入環境變數
load_dotenv()

class AIProvider(Enum):
    GROK = "GROK"
    FREE_CHATGPT = "FREE_CHATGPT"
    GEMINI = "GEMINI"

class AISettings:
    @staticmethod
    def get_api_settings(provider: AIProvider):
        settings = {
            AIProvider.GROK: {
                "api_key": os.getenv("GROK_API_KEY"),
                "api_url": "https://api.x.ai/v1",
                "model": "grok-2-vision-1212",
            },
            AIProvider.FREE_CHATGPT: {
                "api_key": os.getenv("FREE_CHATGPT_API_KEY"),
                "api_url": "https://api.gpt.ge/v1",
                "model": "gpt-3.5-turbo",
            },
            AIProvider.GEMINI: {
                "api_key": os.getenv("GEMINI_API_KEY"),
                "api_url": None,
                "model": "gemini-pro",
            }
        }
        return settings.get(provider)

    @staticmethod
    def test_api_connection(provider: AIProvider):
        """測試API連接是否正常"""
        settings = AISettings.get_api_settings(provider)
        test_prompt = "Hello, this is a test message."

        try:
            if provider == AIProvider.GEMINI:
                import google.generativeai as genai
                genai.configure(api_key=settings["api_key"])
                model = genai.GenerativeModel(settings["model"])
                response = model.generate_content(test_prompt)
                return True, "Gemini API 連接成功"
            else:
                from openai import OpenAI
                client = OpenAI(
                    api_key=settings["api_key"],
                    base_url=settings["api_url"]
                )

                # 根據不同的提供者使用不同的請求格式
                if provider == AIProvider.GROK:
                    response = client.chat.completions.create(
                        model=settings["model"],
                        messages=[
                            {"role": "system", "content": "You are a helpful assistant."},
                            {"role": "user", "content": test_prompt}
                        ],
                        max_tokens=10
                    )
                elif provider == AIProvider.FREE_CHATGPT:
                    response = client.chat.completions.create(
                        model=settings["model"],
                        messages=[{"role": "user", "content": test_prompt}],
                        max_tokens=10,
                        temperature=0.7
                    )

                return True, f"{provider.value} API 連接成功"
        except Exception as e:
            return False, f"{provider.value} API 連接失敗: {str(e)}"