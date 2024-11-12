# Word文件翻譯工具

這是一個使用Python開發的Word文件翻譯工具，可以將Word文件(.docx)翻譯成英文或泰文，同時保留原始文件的格式。使用的前提是必需要有Open AI API Key

## 功能特點

- 支援將Word文件翻譯成英文或泰文
- 自動翻譯文件名稱和表格內容
- 使用OpenAI API進行翻譯
- 保留原始文件的格式（字體、顏色、段落等）
- 簡單的命令行界面
- 支援批量處理多個文件

## 系統需求

- Python 3.8 或更高版本
- Windows 作業系統

## 必要套件安裝

在使用之前，請先安裝以下Python套件：

```bash
pip install python-docx    # 用於處理Word文件
pip install python-dotenv  # 用於處理環境變數
pip install openai        # OpenAI API 客戶端
```

## 安裝步驟

1. 下載程式並解壓到任意資料夾，例如：`D:\WordTranslator\`
2. 在程式資料夾中建立以下結構：

WordTranslator/
│
├── .env # API設定檔
├── input/ # 存放待翻譯的Word文件
├── output/ # 存放翻譯完成的Word文件
└── translate_doc.py # 主程式
```

3. 安裝必要的Python套件：
```bash
pip install python-docx
pip install python-dotenv
pip install openai
```

4. 在程式資料夾中建立 `.env` 檔案，並加入您的API金鑰：
```
OPENAI_API_KEY=你的API金鑰
```

## 快速開始

1. 克隆專案：
```bash
git clone https://github.com/你的用戶名/word-translator.git
cd word-translator
```

2. 安裝依賴：
```bash
pip install -r requirements.txt
```

3. 設定環境變數：
   - 複製 `.env.example` 為 `.env`
   - 在 `.env` 中填入您的 OpenAI API 金鑰
```bash
cp .env.example .env
```

4. 使用方法：
   - 將要翻譯的 .docx 文件放入 `input` 目錄
   - 執行程式：
```bash
python translate_doc.py
```
   - 翻譯後的文件將出現在 `output` 目錄
