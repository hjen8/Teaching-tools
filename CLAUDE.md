# Teaching-tools — 班級工具總專案

## 工作模式
- 進度與最近更動都在 Obsidian：`E:\Brain\Teaching-tools\工作筆記.md`
- **加新工具**：在 `E:\OpenCode\tools\<工具名>\` 下建立
- **收工 SOP**：
  1. 歸納本次工作內容，建議一個資料夾名稱（例：「PPT教材 自動比對 更新」）
  2. 等用戶同意後，在 `E:\OpenCode\archive\` 下建立該資料夾
  3. 將本次工作產出的檔案（腳本、報表、設定檔等）移入該資料夾
  4. **原始素材檔**（如原始 PPT、Word 等來源檔案）：
     - 先判斷是否對後續測試、修改腳本有保留價值
     - 若有 → 一併移入 archive 資料夾
     - 若無 → 詢問用戶是否要刪除
  5. **産出結果檔**（如腳本產出的更新版 PPT）：
     - 提醒用戶自行搬到對應的教材目錄，不放入 repo
  6. 自動 commit + push + 更新工作筆記
- **PPT/Word 比對工作流程**：每次交付新舊 PPT + 新 Word 檔，必須產出兩個結果：
  1. **修改前後對照表**（HTML，左右對照格式，左舊右新，不含 Shape/ID 內部資訊）
  2. **新 PPT**（套用文字更新 + 頁碼校正）
- **開工**：讀工作筆記、檢查 git 狀態、建議下一步
- AI 功能用 Gemini 免費 API（環境變數 GEMINI_API_KEY）或本地 Ollama（gemma4:e4b）

## 三個家
- 🐙 **GitHub**：`hjen8/Teaching-tools`（公開）
- 📘 **Obsidian**：`E:\Brain\Teaching-tools\工作筆記.md`
- 🔥 **Firebase**：`my-teaching-tools-4dd75`（asia-east1）

## ⚠️ 雙倉同步規則

工具原始碼放在 `E:\OpenCode\tools\`（GitHub 版），對應的說明筆記在 Obsidian `E:\Brain\Teaching-tools\工作筆記.md`。兩份同步修改，只推 GitHub。

## 工具清單

| 工具名稱 | 位置 | 狀態 | 說明 |
|---------|------|------|------|
| （尚無） | | | |

## 工作注意事項
- 學生資料一律去識別化（只用座號 + 班級代號）
- commit 訊息寫清楚做了什麼 + 為什麼
- 收工前說「收工」

## 本日教訓（2026-05-20）
- **頁碼偏移量**：不要硬編碼（如 195），改為從 PPT 動態偵測首個頁碼作為基準
- **HTML 報表格式**：左右對照（左舊右新），移除 Shape/ID 等內部物件名稱
- **文字檔輸出**：用 Python 的 `open(path, 'w', encoding='utf-8')` 寫入，不用 PowerShell `>` redirect（會變 UTF-16 LE 亂碼）
- **檔案命名**：對照表統一命名為「修改前後對照表.html」
