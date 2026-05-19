"""探測 Ch7 PPT 結構"""
import sys
import win32com.client
import os

sys.stdout.reconfigure(encoding='utf-8')

ppt_path = r"E:\OpenCode\Ch7  臺灣與世界.pptx"
app = win32com.client.Dispatch("PowerPoint.Application")
app.Visible = True
pres = app.Presentations.Open(os.path.abspath(ppt_path))

print(f"Slide 數: {pres.Slides.Count}")
print()

for slide_idx in range(1, min(pres.Slides.Count + 1, 40)):
    slide = pres.Slides(slide_idx)
    title_text = ''
    for shape in slide.Shapes:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            tr = shape.TextFrame.TextRange
            text = tr.Text.strip()
            if text and len(text) < 60:
                title_text = text
                break
    print(f"Slide {slide_idx:2d}: {title_text}")

pres.Close()
app.Quit()
