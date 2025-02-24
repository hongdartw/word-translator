# Word文件翻譯工具

這是一個使用Python開發的Word文件翻譯工具，可以將Word文件(.docx)翻譯成英文或泰文，同時保留原始文件的格式。使用的前提是必需要有Open AI API Key

## 功能特點

- 支援多種 AI 翻譯服務（Grok、Free ChatGPT、Gemini）
- 支援將Word文件翻譯成英文或泰文
- 自動翻譯文件名稱和表格內容
- 保留原始文件的格式（字體、顏色、段落等）
- 簡單的命令行界面
- 支援批量處理多個文件
- 支援 ESC 鍵隨時中斷程序
- 自動測試 API 可用性

## 系統需求

- Python 3.8 或更高版本
- Windows 作業系統

## 必要套件安裝

在使用之前，請先安裝以下Python套件：

## 安裝步驟

1. 下載程式並解壓到任意資料夾，例如：`D:\WordTranslator\`
2. 在程式資料夾中建立以下結構：
WordTranslator/
│
├── .env # API設定檔
├── input/ # 存放待翻譯的Word文件
├── output/ # 存放翻譯完成的Word文件
├── ai_settings.py # AI 服務設定
└── translate_doc.py # 主程式

3. 在程式資料夾中建立 `.env` 檔案，並加入您的API金鑰：

```
# Grok API 設定
GROK_API_KEY=your_grok_api_key
GROK_API_URL=https://api.x.ai/v1

# Free ChatGPT API 設定
FREE_CHATGPT_API_KEY=your_free_chatgpt_api_key
FREE_CHATGPT_API_URL=https://api.gpt.ge/v1

# Gemini API 設定
GEMINI_API_KEY=your_gemini_api_key
```

## 使用方法

1. 將要翻譯的 .docx 文件放入 `input` 目錄

2. 執行程式：
```bash
python translate_doc.py
```

3. 程式執行流程：
   - 自動測試所有可用的 API 服務
   - 顯示可用的 API 服務列表供選擇
   - 選擇目標語言（英文或泰文）
   - 開始翻譯文件
   - 翻譯後的文件將出現在 `output` 目錄

4. 隨時可按 ESC 鍵中斷程序

## 支援的 AI 服務

1. Grok (x.ai)
   - 模型：grok-2-vision-1212
   - 需要 x.ai 的 API 金鑰

2. Free ChatGPT
   - 模型：gpt-3.5-turbo
   - 需要 ChatAnywhere 的 API 金鑰

3. Google Gemini
   - 模型：gemini-pro
   - 需要 Google AI Studio 的 API 金鑰

## 注意事項

- 請確保您有相應服務的有效 API 金鑰
- 程式會自動測試 API 的可用性，只顯示可用的服務
- 翻譯過程中可隨時按 ESC 鍵中斷程序
- 建議先使用小文件測試各個 API 的效果

