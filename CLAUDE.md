# Teaching-tools — 班級工具總專案

## 工作模式
- 進度與最近更動都在 Obsidian：`E:\Brain\Teaching-tools\工作筆記.md`
- **加新工具**：在 `E:\OpenCode\tools\<工具名>\` 下建立
- **收工 SOP**：
  1. 歸納本次工作內容，建議一個資料夾名稱（例：「PPT教材 自動比對 更新」）
  2. 等用戶同意後，在 `E:\OpenCode\archive\` 下建立該資料夾
  3. 將本次工作產出的檔案（腳本、報表、設定檔等）移入該資料夾
  4. 自動 commit + push + 更新工作筆記
- **開工**：讀工作筆記、檢查 git 狀態、建議下一步
- AI 功能用 Gemini 免費 API（環境變數 GEMINI_API_KEY）或本地 Ollama（gemma4:e4b）

## 三個家
- 🐙 **GitHub**：`hjen8/Teaching-tools`（公開）
- 📘 **Obsidian**：`E:\Brain\Teaching-tools\工作筆記.md`
- 🔥 **Firebase**：`my-teaching-tools-4dd75`（asia-east1）

## ⚠️ 雙倉同步規則

工具原始碼放在 `E:\OpenCode\tools\`（GitHub 版），對應的說明筆記在 Obsidian `E:\Brain\Teaching-tools\`。兩份同步修改，只推 GitHub。

## 工具清單

| 工具名稱 | 位置 | 狀態 | 說明 |
|---------|------|------|------|
| （尚無） | | | |

## 工作注意事項
- 學生資料一律去識別化（只用座號 + 班級代號）
- commit 訊息寫清楚做了什麼 + 為什麼
- 收工前說「收工」
