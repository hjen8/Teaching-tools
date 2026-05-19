"""探測 Ch7 PPT 中的頁碼"""
import sys
import win32com.client
import os
import re

sys.stdout.reconfigure(encoding='utf-8')

ppt_path = r"E:\OpenCode\Ch7  臺灣與世界.pptx"
app = win32com.client.Dispatch("PowerPoint.Application")
app.Visible = True
pres = app.Presentations.Open(os.path.abspath(ppt_path))

print("=== Ch7 PPT 頁碼清單 ===")
page_nums = {}
for slide_idx in range(1, pres.Slides.Count + 1):
    slide = pres.Slides(slide_idx)
    for shape in slide.Shapes:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            tr = shape.TextFrame.TextRange
            text = tr.Text
            for m in re.finditer(r'([pP]\.)(\d+)', text):
                num = int(m.group(2))
                if num not in page_nums:
                    page_nums[num] = []
                page_nums[num].append(slide_idx)

for num in sorted(page_nums.keys()):
    slides = page_nums[num]
    print(f"  p.{num} → Slide {slides}")

pres.Close()
app.Quit()
