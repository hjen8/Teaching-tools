# OpenCode 版懶人包總備份

> 原始來源：Claude Code 懶人包 00~08（mathruffian-dot/claude-code-lazy-packs）
> 轉換日期：2026-05-18
> 對應工具：OpenCode（非 Claude Code）
> 使用者：數學／地理／科學老師

---

## 一、轉換摘要

| 原始項目 | Claude Code 版 | → | OpenCode 版 |
|---------|---------------|----|------------|
| MCP 設定 | `.claude/settings.json` | → | `opencode.json` 的 `mcp` 區塊 |
| Slash 指令 | `/收工` `/開工` skills | → | `opencode.json` 的 `agents` + `instructions` |
| Skill 目錄 | `~/.claude-skills/` | → | 無（改用 agent 指令約定） |
| 專案藍圖 | 各專案 `CLAUDE.md` | → | 同左（OpenCode 也支援） |
| API Key | 環境變數 + `.claude/` | → | Windows 環境變數（`GEMINI_API_KEY`） |
| 安裝工具 | `uv`, `pip`, `npm` | → | 同左（平台無關） |

---

## 二、檔案清單

### 核心設定（E:\OpenCode\）

| 檔案 | 來自懶人包 | 說明 |
|------|-----------|------|
| `opencode.json` | 03, 04.5, 06, 07 | MCP 伺服器（notebooklm/obsidian/firebase）+ 收工/開工 agent |
| `CLAUDE.md` | 07 | 專案藍圖（三個家 + 雙倉同步 + 工具清單表格） |
| `.gitignore` | 07 | 排除敏感檔、系統檔、node_modules |
| `firestore.rules` | 04.5 | Firestore 白名單規則 |
| `.firebaserc` | 04.5 | Firebase 專案 ID |
| `firebase.json` | 04.5 | Firebase CLI 設定 |
| `tools/README.md` | 07 | 工具一覽表格模板 |

### 第二大腦（E:\Brain\）

| 檔案 | 來自懶人包 | 說明 |
|------|-----------|------|
| `CLAUDE.md` | 03 | Vault 工作規則（三層結構 + Clippings 規則） |
| `Teaching-tools/工作筆記.md` | 07 | 進度駕駛艙 |
| `知識庫/index.md` | 03+ | 知識庫目錄索引 |
| `知識庫/log.md` | 03+ | 知識重整操作紀錄 |
| `Templates/每日筆記.md` | 03+ | 每日筆記模板 |
| `Templates/週計畫.md` | 03+ | 週計畫模板 |
| `Templates/知識庫頁面.md` | 03+ | 知識庫頁面模板 |

### 網頁成品（E:\NotebookLM\）

| 檔案 | 來自懶人包 | 說明 |
|------|-----------|------|
| `wordcloud-demo/index.html` | 04.5 | 即時文字雲（Firebase onSnapshot） |
| `gemini-test/index.html` | 06 | Gemini API 測試頁 |

### 系統環境

| 項目 | 來自懶人包 |
|------|-----------|
| Node.js v24.15.0 | 00 |
| npm 11.12.1 | 00 |
| Git v2.52.0 | 00 |
| GitHub CLI（已登入 hjen8） | 00 |
| Ollama v0.24.0 + gemma4:e4b（9.6 GB） | 05 |
| 環境變數 `GEMINI_API_KEY` | 06 |
| Google Drive 鏡射同步 | 03 |
| Obsidian Vault `E:\Brain` | 03 |
| Firebase 專案 `my-teaching-tools-4dd75` | 04.5 |
| GitHub repo `hjen8/Teaching-tools`（公開） | 07 |

---

## 三、完整檔案內容

### 3.1 opencode.json

```json
{
  "$schema": "https://opencode.ai/config.json",
  "instructions": [
    "請將所有 NotebookLM 的成品下載到 E:\\NotebookLM 及其子資料夾中。",
    "我的 Obsidian 第二大腦 Vault 位於 E:\\Brain，使用 mcpvault (obsidian MCP) 存取。",
    "我是數學、地理、科學老師，所有回應請使用繁體中文。",
    "所有筆記、檔案名稱一律使用繁體中文，禁用簡體中文。",
    "Firebase 專案 ID 為 my-teaching-tools-4dd75，透過 firebase MCP 存取 Cloud Firestore。",
    "班級工具原始碼放在 E:\\OpenCode\\tools\\ 內，GitHub repo 為 hjen8/Teaching-tools。",
    "當我說「收工」時，請執行：1) 摘要今天變動 2) 更新 E:\\Brain\\Teaching-tools\\工作筆記.md 的「上次做到哪」和「最近更動紀錄」 3) git add + commit + push。",
    "當我說「開工」時，請執行：1) 讀 E:\\Brain\\Teaching-tools\\工作筆記.md 摘要進度 2) git status 檢查本地狀態 3) git fetch 檢查遠端狀態 4) 建議下一步。"
  ],
  "mcp": {
    "notebooklm-mcp": {
      "type": "local",
      "command": "notebooklm-mcp",
      "enabled": true
    },
    "obsidian": {
      "type": "local",
      "command": "C:\\Users\\R\\AppData\\Roaming\\npm\\mcpvault.cmd",
      "args": ["E:\\Brain"],
      "enabled": true
    },
    "firebase": {
      "type": "local",
      "command": "firebase",
      "args": ["mcp"],
      "enabled": true
    }
  },
  "agents": {
    "shutdown": {
      "description": "收工同步：commit + push + 更新 Obsidian 工作筆記",
      "prompt": "你現在是收工同步助手。請依序執行：\n1. 從對話歷史摘要今天完成的檔案與決策\n2. 讀取工作筆記 E:\\Brain\\Teaching-tools\\工作筆記.md\n3. 更新工作筆記：\n   - ⏯️「上次做到哪」段：更新最後動作、完成的檔案、對話脈絡\n   - 🗓️「最近更動紀錄」表格最後加一行：日期 + 摘要 + ✅✅\n   - 🕳️「踩坑筆記」若有新坑加進去\n4. git add + git commit（commit message 寫清楚做了什麼 + 為什麼）\n5. git push\n6. 回報三欄表格：GitHub + Obsidian 同步狀態\n\n注意：\n- 沒有實質進度時不要跑同步\n- commit message 要有資訊，不要只寫「更新」"
    },
    "startup": {
      "description": "開工接續：讀工作筆記、回報進度、建議下一步",
      "prompt": "你現在是開工接續助手。請依序執行：\n1. 讀取工作筆記 E:\\Brain\\Teaching-tools\\工作筆記.md\n2. 摘要「上次做到哪」（精簡 1-2 句，不要全文倒出）\n3. 執行 git status --short 檢查本地狀態\n4. 執行 git fetch origin 檢查遠端（若 30 分鐘內 fetch 過可跳過）\n5. 檢查落後：git rev-list HEAD..origin/HEAD --count\n6. 回報結構化摘要：\n   📂 專案：Teaching-tools\n   📘 上次進度：（摘要）\n   🔧 本地：（clean / N 個未 commit）\n   🌐 遠端：（最新 / 落後 N commits）\n   ➡️ 建議下一步：（列出可選方向）\n\n注意：\n- 不主動 git pull\n- 不主動修改工作筆記\n- 用摘要方式呈現，不要全文倒出"
    }
  }
}
```

### 3.2 CLAUDE.md（E:\OpenCode\）

```markdown
# Teaching-tools — 班級工具總專案

## 工作模式
- 進度與最近更動都在 Obsidian：`E:\Brain\Teaching-tools\工作筆記.md`
- **加新工具**：在 `E:\OpenCode\tools\<工具名>\` 下建立
- **收工**：自動 commit + push + 更新工作筆記
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
```

### 3.3 .gitignore

```
# Google Drive 系統檔
desktop.ini
*.tmp
~$*

# 敏感資料
.env
*.key
credentials.*

# Node.js
node_modules/
package-lock.json

# 系統
.DS_Store
Thumbs.db
```

### 3.4 tools/README.md

```markdown
# 班級工具一覽

| 工具名稱 | 位置 | 狀態 | 版本 | 說明 |
|---------|------|------|------|------|
| （尚無） | | | | |
```

### 3.5 firestore.rules

```javascript
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {

    // 白名單：在這裡開放需要的集合
    match /wordcloud_words/{document} {
      allow read, write: if true;
    }

    // 其他集合預設禁止
    match /{document=**} {
      allow read, write: if false;
    }
  }
}
```

### 3.6 .firebaserc

```json
{
  "projects": {
    "default": "my-teaching-tools-4dd75"
  }
}
```

### 3.7 firebase.json

```json
{
  "firestore": {
    "rules": "firestore.rules"
  }
}
```

### 3.8 CLAUDE.md（E:\Brain\ — Obsidian Vault 版）

```markdown
# Brain — 我的第二大腦

## 關於我
- 我是數學、地理、科學老師
  - 國中數學：七、八、九年級
  - 高中數學：高一、二、三
  - 高中地理：高一、二、三
- 這個 vault 是我的教學第二大腦

## 語言偏好
- 所有回應請使用繁體中文
- 所有筆記、檔案名稱一律使用繁體中文，禁用簡體中文

## 筆記庫結構（三層）

| 資料夾 | 用途 | 存什麼 |
|---|---|---|
| `Clippings/` | 輸入 | 網路上剪藏的文章、影片筆記、別人的資料 |
| `知識庫/` | 消化 | AI 整理後的結構化知識（有 index.md 索引） |
| `創作庫/` | 輸出 | 我自己的教案、教材、報告、原創內容 |
| `教學素材/` | 素材 | 教學相關素材與資源 |
| `影片筆記/` | 影片 | 影片相關的筆記與紀錄 |
| `每日筆記/` | 時間管理 | 每日紀錄、週計畫 |
| `Templates/` | 模板 | 各種筆記的固定格式 |

## 工作規則

### 新增筆記時
- 一律加上 frontmatter（title、date、tags）
- 根據內容判斷存到正確的資料夾：
  - 別人的文章 → `Clippings/`
  - AI 整理的知識 → `知識庫/`
  - 自己的作品 → `創作庫/`
  - 教學資源 → `教學素材/`
  - 影片紀錄 → `影片筆記/`
  - 每日/每週紀錄 → `每日筆記/`

### Clippings 處理規則
- 透過 Web Clipper 剪藏的文章自動存到 `Clippings/`
- 這些是原始資料，不要修改
- 每週知識重整時，把 Clippings 消化成知識庫頁面

### 知識庫維護規則
- 知識庫有 `index.md`（目錄）和 `log.md`（操作紀錄）
- 每次新增知識庫頁面都要更新 index.md
- 每次做知識重整都要在 log.md 新增紀錄

### 使用者說「幫我新增到筆記」時
- 代表存到這個 Obsidian vault
- 根據內容判斷存到正確的資料夾
```

### 3.9 工作筆記.md（E:\Brain\Teaching-tools\）

```markdown
# Teaching-tools 工作筆記

> 進度日誌（變動快）。專案藍圖請看 OpenCode 端的 `CLAUDE.md`。
> 進度只在這裡記錄，避免雙寫漂移。

## ⏯️ 上次做到哪

**最後動作**：完成懶人包 #07 班級工具工作模式初始化（OpenCode 版）
**所在 repo**：[hjen8/Teaching-tools](https://github.com/hjen8/Teaching-tools)
**已就緒**：Gemini API Key + Ollama 本地 AI + Firebase 資料庫，可開始製作工具

## 🛠️ 工具清單

| 工具名稱 | 位置 | 狀態 | 說明 |
|---------|------|------|------|
| （尚無） | | | |

## 🗓️ 最近更動紀錄

| 日期 | 變更摘要 | Obsidian | GitHub |
|------|----------|----------|--------|
| 2026-05-18 | 初始化班級工具總專案（.gitignore + tools/） | ✅ | ✅ |
| 2026-05-18 | 初始化班級工具工作模式（opencode.json agents + CLAUDE.md + GitHub） | ✅ | ✅ |

## 🕳️ 踩坑筆記

（之後遇到坑就記在這）
```

### 3.10 Obsidian 模板

**每日筆記（Templates/每日筆記.md）**：
```markdown
---
date: {{date:YYYY-MM-DD}}
type: 每日筆記
tags:
  - 每日筆記
---

# {{title}}

---

## 🏫 教學

-

**備課進度**：

---

## 💡 今日反思

>

---

## 明日優先事項

1.
2.
3.
```

**週計畫（Templates/週計畫.md）**：
```markdown
---
week: {{date:YYYY-[W]WW}}
type: 週計畫
tags:
  - 週計畫
---

# {{title}}

---

## 🏫 教學重點

| 日期 | 班級 | 單元 |
|------|------|------|
| | | |

---

## 📧 本週重要事項

- [ ]

---

## 本週每日連結

- [[{{monday:MM/DD}}（一）]]
- [[{{tuesday:MM/DD}}（二）]]
- [[{{wednesday:MM/DD}}（三）]]
- [[{{thursday:MM/DD}}（四）]]
- [[{{friday:MM/DD}}（五）]]
```

**知識庫頁面（Templates/知識庫頁面.md）**：
```markdown
---
title: ""
type: 知識庫
source: ""
created: {{date:YYYY-MM-DD}}
tags: []
related: []
---

#

> 原始資料：
> 整理日期：{{date:YYYY-MM-DD}}

---

## 核心概念



---

## 與我的教學的連結



---

## 內容缺口（待補）

- [ ]
```

### 3.11 Gemini API 測試頁（E:\NotebookLM\gemini-test\index.html）

```html
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Gemini API 測試工具</title>
<style>
  body { font-family: sans-serif; max-width: 700px; margin: 40px auto; padding: 0 20px; }
  textarea { width: 100%; min-height: 80px; padding: 10px; font-size: 16px; }
  button { padding: 10px 24px; font-size: 16px; cursor: pointer; }
  #response { margin-top: 20px; padding: 15px; background: #f5f5f5; border-radius: 8px; white-space: pre-wrap; }
  .error { color: red; }
</style>
</head>
<body>
<h1>Gemini API 測試工具</h1>
<p>測試 Gemini 2.5 Flash 是否正常運作</p>
<textarea id="prompt" placeholder="請輸入想問 Gemini 的問題…">用繁體中文回答：誰發明了微積分？</textarea>
<br><br>
<button onclick="askGemini()">送出問題</button>
<div id="response"></div>

<script>
const API_KEY = "AIzaSyDhX5Z5n03tWQRhcaJ-00Kql9V-PolFYss";

async function askGemini() {
  const prompt = document.getElementById("prompt").value;
  const respDiv = document.getElementById("response");
  respDiv.innerHTML = "思考中…";
  respDiv.className = "";
  try {
    const res = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`,
      { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }) }
    );
    const data = await res.json();
    if (data.error) { respDiv.innerHTML = `<span class="error">錯誤：${data.error.message}</span>`; return; }
    respDiv.textContent = data.candidates[0].content.parts[0].text;
  } catch (e) {
    respDiv.innerHTML = `<span class="error">連線錯誤：${e.message}</span>`;
  }
}
</script>
</body>
</html>
```

### 3.12 即時文字雲（E:\NotebookLM\wordcloud-demo\index.html）

> 略（186 行，完整 HTML + Firebase onSnapshot 即時同步）
> 如需完整檔案內容請見 `E:\NotebookLM\wordcloud-demo\index.html`

---

### 踩坑紀錄：OpenCode 不支援 `agents` 欄位

原始 `opencode.json` 中加入了 `agents`（shutdown/startup），但 OpenCode 的 JSON schema 不認得此欄位，導致啟動時報 `ConfigInvalidError`。已移除 `agents`，收工/開工邏輯全數保留在 `instructions` 陣列中，運作正常。

---

## 四、各懶人包轉換備註

| 編號 | 原始懶人包 | 轉換為 OpenCode 版說明 | 跳過原因 |
|------|-----------|----------------------|---------|
| 00 | 環境建置 | Node.js/Git/GitHub CLI 皆已安裝，沿用既有環境 | |
| 01 | NotebookLM MCP | opencode.json 已設 notebooklm-mcp | |
| 02 | GitHub | 已連 hjen8，已推 Teaching-tools repo | |
| 03 | Obsidian 第二大腦 | Vault 在 E:\Brain，mcpvault 全域安裝，CLAUDE.md 已建立 | |
| 03+ | 第二大腦設定指南 | 三層結構已建立，模板已建立，Web Clipper 已設 | |
| 04 | Supabase | — | 選擇 Firebase（更適合老師） |
| 04.5 | Firebase | 專案已建，白名單 rules 已部署，文字雲 demo 已做 | |
| 05 | Ollama 本地 AI | gemma4:e4b 已安裝在 E:，REST API 正常 | |
| 06 | Gemini 免費 API | GEMINI_API_KEY 已存入環境變數，測試頁已建立 | |
| 07 | 班級工具工作模式 | opencode.json agents + CLAUDE.md + GitHub repo + 工作筆記 | |
| 08 | gpt-image-2 生圖 | — | Gemini 生圖已夠用，不須花錢 |

---

*本備份檔案內容與實際檔案同步。如需還原，將以上檔案放回對應路徑即可恢復完整工作環境。*
