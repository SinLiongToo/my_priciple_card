---
name: workplace-principles-workflow
description: >-
  Automates the full lifecycle of the workplace principles handbook and card system: parsing PowerPoint slides (.pptx), caching trilingual translations (Mandarin, English, Taiwanese Hokkien), regenerating the Markdown handbook, ODT publication book, and RWD interactive card deck web application (index.html), and deploying updates to GitHub Pages. Use when the user asks to update slides, add or modify workplace principles, customize the card web app, or deploy changes.
---

# Workplace Principles Handbook & Card Deck Workflow

本技能提供「工作的管見」系統的端到端操作手冊與維護指南，涵蓋簡報解析、三語翻譯、手冊/書籍/網頁生成以及 GitHub Pages 自動發布。

---

## 一、系統架構與檔案對照

```text
[工作原則-書.pptx.pptx] (簡報源檔)
         │
         ▼
[update_handbook.py] ◄───► [translation_db.json] (快取資料庫)
         │                         ▲
         │ (呼叫 Gemini 2.5 Flash API 翻譯新項目)
         │
         ├───► [工作原則整理與潤稿.md] (三語手冊)
         ├───► [工作原則整理與潤稿.odt] (實體印刷書籍)
         └───► [index.html] (RWD 自包含三語卡牌 Web App)
                     │
                     ▼
           [GitHub Pages 線上站] (https://sinliongtoo.github.io/my_priciple_card/)
```

---

## 二、標準維護流程 (Standard Operating Procedure)

### 步驟 1：簡報更新或原則異動
當使用者更新了 `工作原則-書.pptx.pptx`：
- 若僅調整格式或文字次序，不需要 API 金鑰。
- 若新增原則或大幅修改英文原文，系統會需要翻譯新項目。確保專案根目錄 `.env` 含有：
  ```env
  GEMINI_API_KEY=your_gemini_api_key_here
  ```

### 步驟 2：執行自動化生成腳本
在專案根目錄執行：
```bash
python update_handbook.py
```
**執行檢查點**：
- 檢查終端輸出中是否成功解析所有項目（如 `Found 458 items in PPTX`）。
- 確認 Step 5 (`.md`)、Step 6 (`index.html`)、Step 7 (`.odt`) 均順利生成且結束代碼為 0。

### 步驟 3：修改網頁卡牌 UI 或新增功能規範
若需要修改網頁卡牌 (`index.html`) 的版面、CSS 樣式或功能（例如新增欄位、修改按鈕、調整 RWD）：
> **重要準則**：
> 1. 永遠在 [`update_handbook.py`](file:///c:/Users/tu-hs/OneDrive/%E6%96%87%E4%BB%B6/2022_0308_MASA/2022-0708/Projects_antigravity/%E5%B7%A5%E4%BD%9C%E7%9A%84%E7%AE%A1%E8%A6%8B-%E6%9B%B8/update_handbook.py) 內的 `html_content` f-string 模板中進行修改。
> 2. 在 Python f-string 內，所有 CSS 與 JavaScript 的花括號必須雙寫為 `{{` 與 `}}`。
> 3. 修改完成後，重新執行 `python update_handbook.py` 重新生成 `index.html`，以保持程式碼的一致性。

### 步驟 4：推送更新至 GitHub Pages
當手冊或網頁更新完成後，執行以下指令推送到遠端：
```bash
git add .
git commit -m "feat: 更新職場原則內容與網頁卡牌"
git push origin main
```
推送完成後約 1 分鐘，[GitHub Pages 線上網站](https://sinliongtoo.github.io/my_priciple_card/) 將自動完成構建並更新。

---

## 三、三語翻譯標準與要求

所有翻譯項目必須恪守以下規範：
1. **優化中文 (國語 / Mandarin)**：
   - 語氣穩重、專業、富有商業智慧。
   - 解決痛點、提供具體職場防線或行動建議。
2. **優化英文 (Polished English)**：
   - 使用母語者道地的現代商務英語，語法精練無贅字。
3. **優化台文 (Taiwanese Hokkien)**：
   - 嚴格使用**標準台語漢字**（避免台羅音標與亂碼注音）。
   - 常用詞例：`鬥相共` (幫忙)、`省氣力` (省力)、`做代誌` (做事情)、`攰` (累)、`毋通` (不要)、`冗雜` (繁雜)、`歇睏` (休息)。

---

## 四、常見問題排查 (Troubleshooting)

### 1. 新增的項目沒有翻譯或顯示 `[TBD]`
- **原因**：環境中缺少 `GEMINI_API_KEY` 或 API 呼叫受限。
- **解法**：檢查根目錄 `.env` 檔案，確保金鑰有效。或手動將高品質翻譯補入 `translation_db.json` 後重新執行 `python update_handbook.py`。

### 2. 投影片編號跳號或子項目被拆開
- **原因**：簡報在同一張 Slide 中使用了多個不同大小的文字框。
- **解法**：`update_handbook.py` 內建啟發式判斷 `current_num - num > 20` 即判定為子項目並合併。如有特例，可微調 `parse_pptx_items` 中的正則或門檻值。

### 3. GitHub Pages 樣式跑版或載入失敗
- **檢查點**：
  - 確保根目錄存在 `.nojekyll`。
  - 確保 GitHub 倉庫 Settings -> Pages 中的 Source 為 `Deploy from a branch` (Branch: `main` / `/ (root)`).
  - 確保所有靜態資源與字體均使用 HTTPS 外部 CDN 連結。
