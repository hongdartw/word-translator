import os
from docx import Document
from openai import OpenAI
from dotenv import load_dotenv
from pathlib import Path

# 修改環境變數和路徑設定
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(BASE_DIR, '.env')
INPUT_DIR = os.path.join(BASE_DIR, 'input')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

# 載入環境變數
load_dotenv(ENV_PATH)

# 初始化OpenAI客戶端
client = OpenAI(
    api_key=os.getenv('OPENAI_API_KEY'),
    base_url='https://api.gpt.ge/v1'
)

def translate_text(text, target_language):
    """
    使用OpenAI API翻譯文本
    """
    # 檢查文本是否包含URL
    if text.startswith('http://') or text.startswith('https://'):
        return text  # 如果是URL，直接返回原文

    try:
        prompt = f"""
        Translate the following text to {target_language}.
        Important rules:
        1. Keep it simple and direct
        2. Do not add any explanatory text like 'The translation of ... is ...'
        3. Just provide the translation

        Text to translate: {text}
        """

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a translator. Provide direct translations without additional explanations."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=1000
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"翻譯時發生錯誤: {str(e)}")
        return text

def translate_filename(filename, target_language):
    """
    翻譯檔案名稱（不包含副檔名）
    """
    name_without_ext = os.path.splitext(filename)[0]
    translated_name = translate_text(name_without_ext, target_language)
    return f"{translated_name}.docx"

def process_document(input_file, target_language):
    """
    處理Word文件並進行翻譯

    Args:
        input_file (str): 輸入文件路徑
        target_language (str): 目標語言
    """
    try:
        # 讀取文件
        doc = Document(input_file)

        # 翻譯段落
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                translated_text = translate_text(paragraph.text, target_language)
                paragraph.text = translated_text

        # 翻譯表格內容
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        translated_text = translate_text(cell.text, target_language)
                        cell.text = translated_text

        # 翻譯並生成新的檔案名
        original_filename = os.path.basename(input_file)
        translated_filename = translate_filename(original_filename, target_language)

        # 生成輸出文件路徑
        output_file = os.path.join(OUTPUT_DIR, translated_filename)

        # 保存翻譯後的文件
        doc.save(output_file)
        print(f"文件已翻譯並保存至: {output_file}")

    except Exception as e:
        print(f"處理文件時發生錯誤: {str(e)}")

def main():
    """
    主程序
    """
    # 確保輸出目錄存在
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

    # 獲取輸入目錄中的所有.docx文件
    docx_files = [f for f in os.listdir(INPUT_DIR) if f.endswith('.docx')]

    if not docx_files:
        print("在輸入目錄中未找到.docx文件")
        return

    # 顯示可用的文件
    print("可用的文件：")
    for i, file in enumerate(docx_files, 1):
        print(f"{i}. {file}")

    # 選擇目標語言
    print("\n選擇目標語言：")
    print("1. 英文 (English)")
    print("2. 泰文 (Thai)")

    language_choice = input("請輸入選項 (1 或 2): ").strip()
    target_language = "english" if language_choice == "1" else "thai"

    # 處理每個文件
    for file in docx_files:
        input_path = os.path.join(INPUT_DIR, file)
        print(f"\n正在處理文件: {file}")
        process_document(input_path, target_language)

if __name__ == "__main__":
    main()