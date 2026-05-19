"""探測 Ch7 Word 頁碼"""
import sys
import win32com.client
import os
import re

sys.stdout.reconfigure(encoding='utf-8')

doc_path = r"E:\OpenCode\地理(2)學習寶典_08_CH7_臺灣與世界_教用.doc"
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(os.path.abspath(doc_path))

for section in doc.Sections:
    header = section.Headers(1)
    if header.Exists:
        header_text = header.Range.Text.strip()
        print(f"Header: '{header_text}'")
        numbers = re.findall(r'\d+', header_text)
        if numbers:
            print(f"頁碼: {numbers[-1]}")

doc.Close(False)
word.Quit()
