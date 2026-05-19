"""
PPT 教材自動比對與更新工具
============================
Phase 1: 比對去年 PPT 與今年 Word 教材，產出 HTML 差異報表
Phase 2: 根據報表，更新 PPT 內的文字（保留格式、動畫、Shape ID）

用法:
  python compare_ppt_word.py report    # 產出差異報表
  python compare_ppt_word.py apply     # 套用變更到 PPT

依賴: pip install pywin32
"""
import sys
import os
import json
import difflib
import re
from datetime import datetime

# ============================================================
# 全域設定
# ============================================================
PPT_PATH = r"E:\OpenCode\Ch8  東亞文化圈的形成與發展.pptx"
WORD_PATH = r"E:\OpenCode\地理(2)學習寶典_09_CH8_東亞文化圈的形成與發展_教用.doc"
REPORT_PATH = r"E:\OpenCode\diff_report.html"
APPLY_PPT_PATH = r"E:\OpenCode\Ch8_東亞文化圈_更新版.pptx"
MAPPING_PATH = r"E:\OpenCode\section_mapping.json"
MANUAL_MAPPING_PATH = r"E:\OpenCode\manual_mapping.json"

# 判斷 Word 中哪些 Style 是「章節標題」
HEADING_STYLE_PREFIXES = ["1-標題", "1-重點", "1-摘要"]
# 判斷哪些 Style 是「內容段落」
CONTENT_STYLE_PREFIXES = ["1-標題-條列", "1-標題1.", "1-標題(1)", "1-標題", "1-重點", "1-摘要",
                          "2-補充-延伸", "2-補充-說明", "標準", "內文"]
# 跳過的 Style（選擇題、解析等）
SKIP_STYLE_PREFIXES = ["3-選擇題", "3-選擇題解析", "4-"]


# ============================================================
# 1. DocExtractor — 從 Word .doc 提取結構化內容
# ============================================================
class DocExtractor:
    """用 win32com 開啟 Word .doc，依 Style 分段落提取"""

    def __init__(self, doc_path):
        self.doc_path = doc_path
        self.sections = []  # [{"heading": str, "paragraphs": [str]}]
        self.all_paragraphs = []  # 原始段落清單

    def extract(self):
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(self.doc_path))

        try:
            self._parse_paragraphs(doc)
            self._build_sections()
        finally:
            doc.Close(False)
            word.Quit()

        return self.sections

    def _parse_paragraphs(self, doc):
        """遍歷所有段落，記錄 style + 文字"""
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            style_name = p.Style.NameLocal if p.Style else "(無)"
            text = p.Range.Text.strip()
            font_size = p.Range.Font.Size
            font_bold = bool(p.Range.Font.Bold) if p.Range.Font.Bold != 9999999 else False

            self.all_paragraphs.append({
                "index": i,
                "style": style_name,
                "text": text,
                "font_size": font_size,
                "font_bold": font_bold,
            })

    def _build_sections(self):
        """依標題段落切分章節"""
        current_section = None

        for p in self.all_paragraphs:
            is_heading = self._is_heading(p)
            is_skip = self._is_skip(p)

            if is_skip:
                continue

            if is_heading and p["text"]:
                # 新章節
                if current_section:
                    self.sections.append(current_section)
                current_section = {
                    "heading": p["text"],
                    "paragraphs": [],
                }
            elif current_section and p["text"]:
                current_section["paragraphs"].append(p["text"])

        if current_section:
            self.sections.append(current_section)

    def _is_heading(self, p):
        """判斷是否為章節標題"""
        style = p["style"]
        for prefix in HEADING_STYLE_PREFIXES:
            if style.startswith(prefix):
                return True
        # Fallback: 字型 >= 14pt 且粗體
        if p["font_size"] >= 14 and p["font_bold"]:
            return True
        return False

    def _is_skip(self, p):
        """判斷是否要跳過（選擇題等）"""
        style = p["style"]
        for prefix in SKIP_STYLE_PREFIXES:
            if style.startswith(prefix):
                return True
        return False


# ============================================================
# 2. PptExtractor — 從 PPT 提取結構化內容（含格式）
# ============================================================
class PptExtractor:
    """用 win32com 開啟 PPT，提取每頁每個 shape 的文字與格式"""

    def __init__(self, ppt_path):
        self.ppt_path = ppt_path
        self.slides = []  # [{"slide_idx": int, "title": str, "shapes": [...]}]

    def extract(self):
        import win32com.client
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True
        pres = app.Presentations.Open(os.path.abspath(self.ppt_path))

        try:
            for slide_idx in range(1, pres.Slides.Count + 1):
                slide = pres.Slides(slide_idx)
                slide_data = self._extract_slide(slide, slide_idx)
                self.slides.append(slide_data)
        finally:
            pres.Close()
            app.Quit()

        return self.slides

    def _extract_slide(self, slide, slide_idx):
        shapes_data = []
        title_text = ""

        for shape in slide.Shapes:
            shape_data = {
                "id": shape.Id,
                "name": shape.Name,
                "type": shape.Type,
                "paragraphs": [],
            }

            if shape.HasTextFrame and shape.TextFrame.HasText:
                tf = shape.TextFrame
                tr = tf.TextRange
                paragraphs = self._extract_textframe(tr)
                shape_data["paragraphs"] = paragraphs

                # 嘗試辨識標題
                if not title_text and paragraphs:
                    first_para_text = paragraphs[0]["text"].strip()
                    if first_para_text and len(first_para_text) < 60:
                        title_text = first_para_text

            shapes_data.append(shape_data)

        return {
            "slide_idx": slide_idx,
            "title": title_text,
            "shapes": shapes_data,
        }

    def _extract_textframe(self, text_range):
        """提取 TextFrame 內所有 paragraph 及其 runs"""
        paragraphs = []
        try:
            para_count = text_range.Paragraphs().Count
        except Exception:
            return paragraphs

        for pi in range(1, para_count + 1):
            try:
                para = text_range.Paragraphs(pi)
                para_text = para.Text.rstrip("\r")

                runs = []
                try:
                    run_count = para.Runs().Count
                    for ri in range(1, run_count + 1):
                        r = para.Runs(ri)
                        runs.append({
                            "text": r.Text,
                            "font_name": r.Font.Name,
                            "font_size": r.Font.Size,
                            "bold": bool(r.Font.Bold) if r.Font.Bold != 9999999 else False,
                            "italic": bool(r.Font.Italic) if r.Font.Italic != 9999999 else False,
                            "color_rgb": r.Font.Color.RGB,
                            "underline": bool(r.Font.Underline) if r.Font.Underline != 9999999 else False,
                        })
                except Exception:
                    # 如果無法讀取 runs， fallback 到整個 paragraph
                    runs = [{
                        "text": para_text,
                        "font_name": para.Font.Name,
                        "font_size": para.Font.Size,
                        "bold": bool(para.Font.Bold) if para.Font.Bold != 9999999 else False,
                        "italic": bool(para.Italic) if hasattr(para, 'Italic') and para.Italic != 9999999 else False,
                        "color_rgb": para.Font.Color.RGB,
                        "underline": False,
                    }]

                paragraphs.append({
                    "text": para_text,
                    "runs": runs,
                })
            except Exception:
                continue

        return paragraphs


# ============================================================
# 3. DiffEngine — 比對 PPT 與 Word 的差異
# ============================================================
class DiffEngine:
    """比對 PPT slides 與 Word sections，找出文字差異"""

    def __init__(self, ppt_slides, word_sections):
        self.ppt_slides = ppt_slides
        self.word_sections = word_sections
        self.mappings = []
        self.changes = []

    def compute(self, manual_mapping_path=None):
        """執行比對"""
        if manual_mapping_path and os.path.exists(manual_mapping_path):
            self._load_manual_mapping(manual_mapping_path)
        else:
            self._map_sections_by_content()
        self._diff_mapped()
        return self.changes

    def _load_manual_mapping(self, path):
        """從手動對照表載入 mapping"""
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

        for m in data["mappings"]:
            word_idx = m["word_idx"]
            slide_range = m["slide_range"]
            section = self.word_sections[word_idx]

            # 收集對應的 slides
            slide_indices = []
            for si in range(slide_range[0] - 1, slide_range[1]):
                if si < len(self.ppt_slides):
                    slide_indices.append(self.ppt_slides[si]["slide_idx"])

            if slide_indices:
                first_slide = self.ppt_slides[slide_range[0] - 1]
                self.mappings.append({
                    "slide_indices": slide_indices,
                    "word_idx": word_idx,
                    "similarity": 1.0,  # 手動對應，視為 100% 準確
                    "slide_range": f"{slide_range[0]}-{slide_range[1]}",
                    "word_heading": section["heading"],
                    "first_slide_title": first_slide["title"],
                })

        print(f"  [手動對應] 載入 {len(self.mappings)} 組映射")

    def _map_sections_by_content(self):
        """用簡化策略：依序將 PPT slides 分組對應到 Word sections"""
        n_slides = len(self.ppt_slides)
        n_sections = len(self.word_sections)

        # 收集每個 slide 的所有文字
        slide_texts = []
        for slide in self.ppt_slides:
            all_text = []
            for shape in slide["shapes"]:
                for para in shape["paragraphs"]:
                    if para["text"].strip():
                        all_text.append(para["text"].strip())
            slide_texts.append(" ".join(all_text))

        # 計算每個 slide 與每個 section 的相似度
        sim_matrix = []
        for si in range(n_slides):
            row = []
            for wi in range(n_sections):
                score = self._content_similarity(slide_texts[si], self.word_sections[wi])
                row.append(score)
            sim_matrix.append(row)

        # 找出每個 section 的「高峰 slide」— 與該 section 相似度最高的 slide
        peak_slides = []
        for wi in range(n_sections):
            best_si = 0
            best_score = 0
            for si in range(n_slides):
                if sim_matrix[si][wi] > best_score:
                    best_score = sim_matrix[si][wi]
                    best_si = si
            peak_slides.append(best_si)

        # 用高峰 slide 作為分界點
        # section i 的 slides 範圍：peak_slides[i] 到 peak_slides[i+1]-1
        slide_to_section = [0] * n_slides
        for wi in range(n_sections):
            start = peak_slides[wi]
            if wi + 1 < n_sections:
                end = peak_slides[wi + 1]
            else:
                end = n_slides

            for si in range(start, min(end, n_slides)):
                slide_to_section[si] = wi

        # 處理 peak_slides 之前的 slides
        if peak_slides[0] > 0:
            for si in range(peak_slides[0]):
                slide_to_section[si] = 0

        # 平滑處理
        slide_to_section = self._smooth_assignments(slide_to_section, n_sections)

        # 建立 mapping
        section_to_slides = {}
        for si, wi in enumerate(slide_to_section):
            if wi not in section_to_slides:
                section_to_slides[wi] = []
            section_to_slides[wi].append(si)

        for wi in range(n_sections):
            if wi in section_to_slides and section_to_slides[wi]:
                slides = section_to_slides[wi]
                section = self.word_sections[wi]
                first_slide = self.ppt_slides[slides[0]]
                avg_sim = sum(sim_matrix[si][wi] for si in slides) / len(slides)

                self.mappings.append({
                    "slide_indices": [self.ppt_slides[si]["slide_idx"] for si in slides],
                    "word_idx": wi,
                    "similarity": avg_sim,
                    "slide_range": f"{self.ppt_slides[slides[0]]['slide_idx']}-{self.ppt_slides[slides[-1]]['slide_idx']}",
                    "word_heading": section["heading"],
                    "first_slide_title": first_slide["title"],
                })

    def _smooth_assignments(self, assignments, n_sections):
        """平滑處理：移除孤立的單頁分配"""
        result = assignments[:]
        changed = True
        while changed:
            changed = False
            for i in range(1, len(result) - 1):
                if result[i] != result[i - 1] and result[i] != result[i + 1]:
                    # 孤立頁，取前後的中間值
                    result[i] = min(result[i - 1], result[i + 1])
                    changed = True
        return result

    def _content_similarity(self, slide_text, word_section):
        """計算 slide 內容與 Word section 的相似度"""
        if not slide_text.strip():
            return 0

        # 提取關鍵字（中文字 + 英文單字）
        slide_words = set(re.findall(r'[\u4e00-\u9fff]+|[a-zA-Z]+', slide_text))
        section_words = set()
        for para in word_section["paragraphs"]:
            section_words.update(re.findall(r'[\u4e00-\u9fff]+|[a-zA-Z]+', para))

        if not slide_words or not section_words:
            return 0

        # Jaccard 相似度
        intersection = slide_words & section_words
        union = slide_words | section_words

        if not union:
            return 0

        jaccard = len(intersection) / len(union)

        # 也計算包含率（slide 中有多少字出現在 section 中）
        coverage = len(intersection) / len(slide_words) if slide_words else 0

        # 綜合分數
        return jaccard * 0.4 + coverage * 0.6

    def _similarity(self, s1, s2):
        """計算兩個字串的相似度（保留舊方法）"""
        s1_clean = re.sub(r'[^\w\u4e00-\u9fff]', '', s1)
        s2_clean = re.sub(r'[^\w\u4e00-\u9fff]', '', s2)

        if not s1_clean or not s2_clean:
            return 0

        if s1_clean in s2_clean or s2_clean in s1_clean:
            return 0.8

        return difflib.SequenceMatcher(None, s1_clean, s2_clean).ratio()

    def _diff_mapped(self):
        """對已對應的 slide-section 做段落 diff"""
        for mapping in self.mappings:
            section = self.word_sections[mapping["word_idx"]]
            word_paras = [p for p in section["paragraphs"] if p.strip()]

            # 收集所有對應 slides 的 paragraph 文字
            ppt_paras = []
            for slide_idx in mapping["slide_indices"]:
                slide = self.ppt_slides[slide_idx - 1]
                for shape in slide["shapes"]:
                    for para in shape["paragraphs"]:
                        if para["text"].strip():
                            ppt_paras.append({
                                "slide_idx": slide_idx,
                                "shape_id": shape["id"],
                                "shape_name": shape["name"],
                                "text": para["text"].strip(),
                                "runs": para["runs"],
                            })

            # 用 SequenceMatcher 做 diff
            matcher = difflib.SequenceMatcher(None,
                                              [p["text"] for p in ppt_paras],
                                              word_paras)

            for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                if tag == "equal":
                    continue

                # 取第一個 affected slide 的資訊
                first_slide_idx = ppt_paras[i1]["slide_idx"] if ppt_paras[i1:i2] else mapping["slide_indices"][0]
                first_slide = self.ppt_slides[first_slide_idx - 1]

                change = {
                    "slide_idx": first_slide_idx,
                    "slide_range": mapping["slide_range"],
                    "slide_title": first_slide["title"],
                    "word_heading": section["heading"],
                    "tag": tag,
                    "ppt_paras": ppt_paras[i1:i2],
                    "word_paras": word_paras[j1:j2],
                    "similarity": mapping["similarity"],
                }

                # 對 replace 做 run-level diff
                if tag == "replace" and i2 - i1 == 1 and j2 - j1 == 1:
                    change["run_diff"] = self._diff_runs(
                        ppt_paras[i1]["runs"],
                        word_paras[j1]
                    )

                self.changes.append(change)

    def _diff_runs(self, ppt_runs, new_text):
        """對 run 層級做 diff，找出哪些 run 需要替換"""
        old_text = "".join(r["text"] for r in ppt_runs)
        matcher = difflib.SequenceMatcher(None, old_text, new_text)

        run_diffs = []
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            run_diffs.append({
                "tag": tag,
                "old_text": old_text[i1:i2],
                "new_text": new_text[j1:j2],
            })

        return run_diffs


# ============================================================
# 4. DiffReporter — 產生 HTML 差異報表
# ============================================================
class DiffReporter:
    """產生 HTML 差異報表"""

    def __init__(self, changes, mappings, ppt_slides, word_sections):
        self.changes = changes
        self.mappings = mappings
        self.ppt_slides = ppt_slides
        self.word_sections = word_sections

    def generate(self, output_path):
        html = self._build_html()
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)
        return output_path

    def _build_html(self):
        stats = self._compute_stats()

        html = f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<title>PPT 教材差異報表</title>
<style>
body {{ font-family: "Microsoft JhengHei", "PingFang TC", sans-serif; margin: 20px; background: #f5f5f5; }}
h1 {{ color: #333; border-bottom: 3px solid #4CAF50; padding-bottom: 10px; }}
h2 {{ color: #555; margin-top: 30px; }}
.stats {{ background: #e8f5e9; padding: 15px; border-radius: 8px; margin: 20px 0; }}
.stats span {{ font-weight: bold; color: #2e7d32; }}
.slide-block {{ background: white; border: 1px solid #ddd; border-radius: 8px; margin: 15px 0; padding: 15px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
.slide-title {{ font-size: 1.2em; font-weight: bold; color: #1565c0; margin-bottom: 10px; }}
.mapping-info {{ font-size: 0.9em; color: #666; margin-bottom: 10px; }}
.change-block {{ margin: 10px 0; padding: 10px; border-left: 4px solid #ccc; background: #fafafa; }}
.change-replace {{ border-left-color: #ff9800; }}
.change-insert {{ border-left-color: #4CAF50; }}
.change-delete {{ border-left-color: #f44336; }}
.old-text {{ background: #ffebee; color: #c62828; padding: 5px 10px; border-radius: 4px; margin: 3px 0; white-space: pre-wrap; }}
.new-text {{ background: #e8f5e9; color: #2e7d32; padding: 5px 10px; border-radius: 4px; margin: 3px 0; white-space: pre-wrap; }}
.shape-info {{ font-size: 0.85em; color: #888; }}
.run-diff {{ background: #fff3e0; padding: 8px; border-radius: 4px; margin: 5px 0; font-family: monospace; font-size: 0.9em; }}
.run-diff .del {{ background: #ffcdd2; text-decoration: line-through; color: #c62828; }}
.run-diff .add {{ background: #c8e6c9; color: #2e7d32; }}
.unmapped {{ background: #fff9c4; border-left-color: #fdd835; }}
table {{ border-collapse: collapse; width: 100%; margin: 10px 0; }}
th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
th {{ background: #4CAF50; color: white; }}
tr:nth-child(even) {{ background: #f9f9f9; }}
</style>
</head>
<body>
<h1>PPT 教材差異報表</h1>
<p>產生時間: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
<p>PPT: {PPT_PATH}</p>
<p>Word: {WORD_PATH}</p>

<div class="stats">
<h2>摘要</h2>
<p>對應章節: <span>{len(self.mappings)}</span> 組</p>
<p>有變動的 slide: <span>{stats["changed_slides"]}</span> 頁</p>
<p>文字修改: <span>{stats["replaces"]}</span> 處</p>
<p>新增段落: <span>{stats["inserts"]}</span> 段</p>
<p>刪除段落: <span>{stats["deletes"]}</span> 段</p>
</div>

<h2>對應關係</h2>
<table>
<tr><th>PPT Slide</th><th>PPT 標題</th><th>Word 章節</th><th>相似度</th></tr>
"""
        for m in self.mappings:
            slide_range = m.get('slide_range', str(m.get('slide_idx', '?')))
            first_title = m.get('first_slide_title', m.get('slide_title', ''))
            html += f"""<tr>
<td>第 {slide_range} 頁</td>
<td>{self._esc(first_title)}</td>
<td>{self._esc(m['word_heading'])}</td>
<td>{m['similarity']:.0%}</td>
</tr>
"""

        html += """</table>

<h2>詳細差異</h2>
"""

        # 按 section 分組顯示
        changes_by_section = {}
        for c in self.changes:
            key = c.get("word_heading", "unknown")
            if key not in changes_by_section:
                changes_by_section[key] = []
            changes_by_section[key].append(c)

        for section_heading in sorted(changes_by_section.keys()):
            section_changes = changes_by_section[section_heading]
            slide_range = section_changes[0].get("slide_range", str(section_changes[0]["slide_idx"]))

            html += f"""<div class="slide-block">
<div class="slide-title">Word: {self._esc(section_heading)} | PPT Slides: {slide_range}</div>
"""
            for change in section_changes:
                tag_class = f"change-{change['tag']}" if change['tag'] in ('replace', 'insert', 'delete') else "change-block"

                html += f'<div class="{tag_class}">'

                if change["tag"] == "replace":
                    html += f'<div class="shape-info">Shape: {self._esc(change["ppt_paras"][0]["shape_name"])} (ID: {change["ppt_paras"][0]["shape_id"]})</div>'
                    for pp in change["ppt_paras"]:
                        html += f'<div class="old-text">舊: {self._esc(pp["text"])}</div>'
                    for wp in change["word_paras"]:
                        html += f'<div class="new-text">新: {self._esc(wp)}</div>'

                    # Run-level diff
                    if change.get("run_diff"):
                        html += '<div class="run-diff">Run 層級差異: '
                        for rd in change["run_diff"]:
                            if rd["tag"] == "delete":
                                html += f'<span class="del">{self._esc(rd["old_text"])}</span>'
                            elif rd["tag"] == "insert":
                                html += f'<span class="add">{self._esc(rd["new_text"])}</span>'
                            elif rd["tag"] == "replace":
                                html += f'<span class="del">{self._esc(rd["old_text"])}</span>'
                                html += f'<span class="add">{self._esc(rd["new_text"])}</span>'
                            else:
                                html += self._esc(rd["old_text"])
                        html += '</div>'

                elif change["tag"] == "insert":
                    html += '<div class="new-text">新增段落:</div>'
                    for wp in change["word_paras"]:
                        html += f'<div class="new-text">{self._esc(wp)}</div>'

                elif change["tag"] == "delete":
                    html += '<div class="old-text">刪除段落:</div>'
                    for pp in change["ppt_paras"]:
                        html += f'<div class="old-text">{self._esc(pp["text"])}</div>'

                html += '</div>'

            html += '</div>'

        # 未對應的 Word 章節
        mapped_word_indices = {m["word_idx"] for m in self.mappings}
        unmapped = [s for i, s in enumerate(self.word_sections) if i not in mapped_word_indices]
        if unmapped:
            html += """<h2>未對應的 Word 章節</h2>"""
            for s in unmapped:
                html += f'<div class="slide-block unmapped"><div class="slide-title">{self._esc(s["heading"])}</div>'
                for p in s["paragraphs"][:3]:
                    html += f'<div class="new-text">{self._esc(p)}</div>'
                if len(s["paragraphs"]) > 3:
                    html += f'<div class="new-text">... 還有 {len(s["paragraphs"]) - 3} 段</div>'
                html += '</div>'

        html += """
</body>
</html>"""
        return html

    def _compute_stats(self):
        changed_slides = set()
        replaces = 0
        inserts = 0
        deletes = 0

        for c in self.changes:
            changed_slides.add(c["slide_idx"])
            if c["tag"] == "replace":
                replaces += 1
            elif c["tag"] == "insert":
                inserts += 1
            elif c["tag"] == "delete":
                deletes += 1

        return {
            "changed_slides": len(changed_slides),
            "replaces": replaces,
            "inserts": inserts,
            "deletes": deletes,
        }

    def _esc(self, text):
        """HTML escape"""
        return (text
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;"))


# ============================================================
# 5. PptUpdater — 執行實際 PPT 文字取代
# ============================================================
class PptUpdater:
    """根據差異報表，更新 PPT 內的文字（保留格式、動畫、Shape ID）"""

    def __init__(self, ppt_path, changes):
        self.ppt_path = ppt_path
        self.changes = changes

    def apply(self, output_path):
        import win32com.client
        import shutil

        # 先複製一份，避免改到原始檔
        shutil.copy2(self.ppt_path, output_path)

        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True
        pres = app.Presentations.Open(os.path.abspath(output_path))

        try:
            # 按 slide 分組變更
            changes_by_slide = {}
            for c in self.changes:
                key = c["slide_idx"]
                if key not in changes_by_slide:
                    changes_by_slide[key] = []
                changes_by_slide[key].append(c)

            for slide_idx in sorted(changes_by_slide.keys()):
                slide = pres.Slides(slide_idx)
                slide_changes = changes_by_slide[slide_idx]
                self._apply_slide_changes(slide, slide_changes)

            pres.Save()
            print(f"[OK] 已儲存至: {output_path}")
        finally:
            pres.Close()
            app.Quit()

    def _apply_slide_changes(self, slide, changes):
        """對單一 slide 套用所有變更"""
        for change in changes:
            if change["tag"] == "replace":
                self._apply_replace(slide, change)
            elif change["tag"] == "insert":
                self._apply_insert(slide, change)
            elif change["tag"] == "delete":
                self._apply_delete(slide, change)

    def _apply_replace(self, slide, change):
        """取代文字（保留 run 格式）"""
        ppt_paras = change["ppt_paras"]
        word_paras = change["word_paras"]

        if not ppt_paras or not word_paras:
            return

        # 找到對應的 shape
        shape_id = ppt_paras[0]["shape_id"]
        shape = self._find_shape_by_id(slide, shape_id)
        if not shape:
            return

        tf = shape.TextFrame
        if not tf.HasText:
            return

        tr = tf.TextRange

        # 策略 1: 如果段落數量相同，逐段取代
        if len(ppt_paras) == len(word_paras):
            for i, (pp, new_text) in enumerate(zip(ppt_paras, word_paras)):
                # 找到對應的 paragraph
                para_idx = self._find_para_index(tr, pp["text"])
                if para_idx is not None:
                    self._replace_paragraph_text(tr.Paragraphs(para_idx), new_text)

        # 策略 2: 段落數量不同，用 run-level diff
        elif change.get("run_diff") and len(ppt_paras) == 1:
            para_idx = self._find_para_index(tr, ppt_paras[0]["text"])
            if para_idx is not None:
                para = tr.Paragraphs(para_idx)
                self._replace_paragraph_text(para, word_paras[0])

    def _apply_insert(self, slide, change):
        """新增段落"""
        ppt_paras = change["ppt_paras"]
        word_paras = change["word_paras"]

        if not ppt_paras or not word_paras:
            return

        shape_id = ppt_paras[0]["shape_id"]
        shape = self._find_shape_by_id(slide, shape_id)
        if not shape:
            return

        tf = shape.TextFrame
        tr = tf.TextRange

        for new_text in word_paras:
            tr.InsertAfter("\r" + new_text)

    def _apply_delete(self, slide, change):
        """刪除段落"""
        ppt_paras = change["ppt_paras"]

        if not ppt_paras:
            return

        shape_id = ppt_paras[0]["shape_id"]
        shape = self._find_shape_by_id(slide, shape_id)
        if not shape:
            return

        tf = shape.TextFrame
        tr = tf.TextRange

        for pp in ppt_paras:
            para_idx = self._find_para_index(tr, pp["text"])
            if para_idx is not None:
                tr.Paragraphs(para_idx).Delete()

    def _find_shape_by_id(self, slide, shape_id):
        """透過 Shape.Id 找到 shape"""
        for shape in slide.Shapes:
            if shape.Id == shape_id:
                return shape
        return None

    def _find_para_index(self, text_range, target_text):
        """在 TextRange 中找到匹配的 paragraph index"""
        try:
            para_count = text_range.Paragraphs().Count
            for i in range(1, para_count + 1):
                para = text_range.Paragraphs(i)
                if para.Text.strip() == target_text.strip():
                    return i
        except Exception:
            pass
        return None

    def _replace_paragraph_text(self, paragraph, new_text):
        """取代 paragraph 文字，保留 run 格式"""
        try:
            # 嘗試保留 run 格式
            run_count = paragraph.Runs().Count
            if run_count > 0:
                # 只改第一個 run 的文字（最安全）
                first_run = paragraph.Runs(1)
                first_run.Text = new_text
                # 刪除其他 runs
                for ri in range(run_count, 1, -1):
                    paragraph.Runs(ri).Delete()
            else:
                paragraph.Text = new_text
        except Exception:
            # Fallback: 直接取代整個 paragraph
            try:
                paragraph.Text = new_text
            except Exception as e:
                print(f"  [WARN] 取代失敗: {e}")


# ============================================================
# 6. main — 流程控制
# ============================================================
def main():
    if len(sys.argv) < 2:
        print("用法:")
        print("  python compare_ppt_word.py report    # 產出差異報表")
        print("  python compare_ppt_word.py apply     # 套用變更到 PPT")
        print("  python compare_ppt_word.py pagenum   # 更新 PPT 頁碼")
        sys.exit(1)

    mode = sys.argv[1]

    if mode == "report":
        run_report()
    elif mode == "apply":
        run_apply()
    elif mode == "pagenum":
        run_pagenum()
    else:
        print(f"未知模式: {mode}")
        sys.exit(1)


def run_report():
    """Phase 1: 產出差異報表"""
    print("=" * 50)
    print("Phase 1: 比對 PPT 與 Word 教材")
    print("=" * 50)

    # 1. 提取 Word 內容
    print("\n[1/4] 提取 Word 內容...")
    doc_ext = DocExtractor(WORD_PATH)
    word_sections = doc_ext.extract()
    print(f"  找到 {len(word_sections)} 個章節")
    for i, s in enumerate(word_sections):
        print(f"    [{i}] {s['heading'][:40]}... ({len(s['paragraphs'])} 段)")

    # 2. 提取 PPT 內容
    print("\n[2/4] 提取 PPT 內容...")
    ppt_ext = PptExtractor(PPT_PATH)
    ppt_slides = ppt_ext.extract()
    print(f"  找到 {len(ppt_slides)} 頁 slide")

    # 3. 比對
    print("\n[3/4] 執行比對...")
    engine = DiffEngine(ppt_slides, word_sections)

    # 檢查是否有手動對照表
    manual_map = MANUAL_MAPPING_PATH if os.path.exists(MANUAL_MAPPING_PATH) else None
    changes = engine.compute(manual_mapping_path=manual_map)
    print(f"  找到 {len(changes)} 處差異")
    print(f"  對應 {len(engine.mappings)} 組 slide-section")

    # 4. 產出報表
    print("\n[4/4] 產出 HTML 報表...")
    reporter = DiffReporter(changes, engine.mappings, ppt_slides, word_sections)
    report_path = reporter.generate(REPORT_PATH)
    print(f"  報表已儲存至: {report_path}")

    # 儲存 mapping 供 apply 使用
    mapping_data = {
        "mappings": engine.mappings,
        "changes": changes,
    }
    with open(MAPPING_PATH, "w", encoding="utf-8") as f:
        json.dump(mapping_data, f, ensure_ascii=False, indent=2)
    print(f"  對應資料已儲存至: {MAPPING_PATH}")

    print("\n完成！請開啟報表檢視差異。")
    print(f"  開啟: {report_path}")


def run_apply():
    """Phase 2: 套用變更到 PPT（文字 + 頁碼）"""
    print("=" * 50)
    print("Phase 2: 套用變更到 PPT")
    print("=" * 50)

    # 讀取 mapping
    if not os.path.exists(MAPPING_PATH):
        print(f"錯誤: 找不到對應資料 {MAPPING_PATH}")
        print("請先執行 report 模式產出對應資料。")
        sys.exit(1)

    with open(MAPPING_PATH, "r", encoding="utf-8") as f:
        mapping_data = json.load(f)

    changes = mapping_data["changes"]
    print(f"\n讀取到 {len(changes)} 處文字變更")

    # Step 1: 套用文字變更
    print("\n[1/2] 套用文字變更...")
    updater = PptUpdater(PPT_PATH, changes)
    updater.apply(APPLY_PPT_PATH)

    # Step 2: 更新頁碼（以文字更新版為基礎）
    print("\n[2/2] 更新頁碼...")
    word_page = get_word_page_number()
    if word_page:
        offset = word_page - 195
        print(f"偏移量: {offset}")
        update_page_numbers(APPLY_PPT_PATH, APPLY_PPT_PATH, offset)
    else:
        print("[WARN] 無法提取 Word 頁碼，跳過頁碼更新")

    print("\n完成！請檢查更新後的 PPT。")
    print(f"  最終版: {APPLY_PPT_PATH}")


def run_pagenum():
    """Phase 3: 更新 PPT 頁碼"""
    print("=" * 50)
    print("Phase 3: 更新 PPT 頁碼")
    print("=" * 50)

    # Step 1: 從 Word 提取頁碼
    print("\n[1/3] 從 Word 提取頁碼...")
    word_page = get_word_page_number()
    if not word_page:
        print("錯誤: 無法從 Word 檔提取頁碼")
        sys.exit(1)

    # Step 2: 找出 PPT 中的頁碼
    print("\n[2/3] 找出 PPT 中的頁碼...")
    ppt_pages = find_ppt_page_numbers(PPT_PATH)
    print(f"找到 {len(ppt_pages)} 處頁碼")

    # 顯示所有找到的頁碼
    old_nums = set()
    for p in ppt_pages:
        old_nums.add(p["old_num"])

    print("\nPPT 中的頁碼:")
    # 計算偏移量：Word 頁碼 - PPT 對應頁碼（假設 p.195 對應 Word 的頁碼）
    offset = word_page - 195
    for num in sorted(old_nums):
        new_num = num + offset
        print(f"  p.{num} → p.{new_num}")

    # Step 3: 更新頁碼
    print("\n[3/3] 更新頁碼...")
    print(f"Word 頁碼: {word_page}")
    print(f"PPT 對應頁碼: 195")
    print(f"偏移量: {offset}")

    update_page_numbers(PPT_PATH, APPLY_PPT_PATH, offset)

    print("\n完成！")


def get_word_page_number():
    """從 Word 檔頁眉提取頁碼"""
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(WORD_PATH))

    page_num = None
    try:
        for section in doc.Sections:
            header = section.Headers(1)  # wdHeaderFooterPrimary
            if header.Exists:
                header_text = header.Range.Text.strip()
                # 搜尋頁碼：通常是最後的數字
                numbers = re.findall(r'\d+', header_text)
                if numbers:
                    page_num = int(numbers[-1])
                    print(f"Word 頁眉: '{header_text}'")
                    print(f"提取頁碼: {page_num}")
                    break
    finally:
        doc.Close(False)
        word.Quit()

    return page_num


def find_ppt_page_numbers(ppt_path):
    """找出 PPT 中所有頁碼"""
    import win32com.client
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True
    pres = app.Presentations.Open(os.path.abspath(ppt_path))

    page_numbers = []

    for slide_idx in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(slide_idx)
        for shape in slide.Shapes:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                tr = shape.TextFrame.TextRange
                text = tr.Text

                # 搜尋頁碼模式
                matches = list(re.finditer(r'([pP]\.)(\d+)', text))
                for match in matches:
                    prefix = match.group(1)
                    old_num = int(match.group(2))

                    page_numbers.append({
                        "slide_idx": slide_idx,
                        "shape_id": shape.Id,
                        "shape_name": shape.Name,
                        "old_num": old_num,
                        "prefix": prefix,
                    })

    pres.Close()
    app.Quit()

    return page_numbers


def update_page_numbers(ppt_path, output_path, offset):
    """更新 PPT 中的所有頁碼"""
    import win32com.client
    import shutil

    # 如果輸入輸出相同，用 SaveAs 到新檔再覆蓋
    same_file = os.path.abspath(ppt_path) == os.path.abspath(output_path)
    if same_file:
        save_path = os.path.join(os.path.dirname(ppt_path), "_temp_save.pptx")
    else:
        shutil.copy2(ppt_path, output_path)
        save_path = output_path

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True
    pres = app.Presentations.Open(os.path.abspath(ppt_path))

    updated_count = 0

    try:
        for slide_idx in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(slide_idx)
            for shape in slide.Shapes:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    tr = shape.TextFrame.TextRange
                    text = tr.Text

                    # 搜尋並替換頁碼（處理範圍格式 p.202~203）
                    def replace_page_num(match):
                        prefix = match.group(1)
                        num1 = int(match.group(2))
                        separator = match.group(3) or ''
                        num2 = int(match.group(4)) if match.group(4) else None

                        new_num1 = num1 + offset
                        if num2:
                            new_num2 = num2 + offset
                            return f"{prefix}{new_num1}{separator}{new_num2}"
                        else:
                            return f"{prefix}{new_num1}"

                    # 匹配模式：p.195 或 p.195~203 或 p.195-203
                    new_text = re.sub(
                        r'([pP]\.)(\d+)([~-]?)(\d*)',
                        replace_page_num,
                        text
                    )

                    if new_text != text:
                        try:
                            tr.Text = new_text
                            updated_count += 1
                            print(f"  Slide {slide_idx}: '{text[:50]}...' → '{new_text[:50]}...'")
                        except Exception as e:
                            print(f"  [WARN] Slide {slide_idx} 更新失敗: {e}")

        pres.SaveAs(os.path.abspath(save_path))
        pres.Close()
        app.Quit()

        # 等 PowerPoint 完全釋放檔案鎖定
        import time
        import gc
        gc.collect()
        time.sleep(2)

        # 覆蓋原檔（重試機制）
        if same_file:
            for _ in range(10):
                try:
                    shutil.copy2(save_path, output_path)
                    os.remove(save_path)
                    break
                except PermissionError:
                    time.sleep(1)
            else:
                print(f"  [WARN] 無法覆蓋原檔，暫存檔保留於: {save_path}")

        print(f"\n[OK] 已更新 {updated_count} 處頁碼")
        print(f"已儲存至: {output_path}")

    except Exception:
        raise
    finally:
        try:
            pres.Close()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
