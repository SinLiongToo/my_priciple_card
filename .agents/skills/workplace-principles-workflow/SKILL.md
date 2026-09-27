---
name: workplace-principles-workflow
description: >-
  Automates the full lifecycle of the workplace principles handbook and card system: parsing PowerPoint slides (.pptx), caching trilingual translations (Mandarin, English, Taiwanese Hokkien), regenerating the Markdown handbook, ODT publication book, and RWD interactive card deck web application (index.html) with Web Speech API audio pronunciation, dual-view continuous autoplay player bars, compact 2-tier responsive header, and relationship connection graph, then deploying updates to GitHub Pages. Use when the user asks to update slides, add or modify workplace principles, customize the card web app, or deploy changes.
---

# Workplace Principles Handbook & Card Deck Workflow

本技能提供「工作的管見」系統的端到端操作手冊與維護指南，涵蓋簡報解析、三語翻譯、手冊/書籍/網頁生成以及 GitHub Pages 自動發布。

---

## 一、系統架構與檔案對照

```text
[工作原則-書.pptx.pptx] (簡報源檔)
         │
         ▼
[update_handbook.py] ◄───► [manual_overrides.json] (👑 最高優先級：人工手動覆蓋庫)
         ▲                         │
         │                         ▼
         │ ◄──────────────► [translation_db.json] (快取資料庫)
         │                         ▲
         │ (僅新項目呼叫 Gemini 3.8 Flash API 翻譯)
         │
          ├───► [工作原則整理與潤稿.md] (三語手冊，含英文書名與副標題)
          ├───► [工作原則整理與潤稿.odt] (實體印刷書籍)
          └───► [index.html] (RWD 自包含三語卡牌 Web App，含 4 大視角與關係圖譜)
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

## 三、手動自訂翻譯覆蓋機制 (Manual Overrides)

專案支援 `manual_overrides.json`，提供獨立維護人工潤飾版本的能力。在此設定的條目權重**永遠高於 `translation_db.json` 與 AI 翻譯**，且腳本更新時**絕不會被覆蓋或沖刷**。

### 使用語法範例 (`manual_overrides.json`)：

```json
{
  "_README": "在此設定的文字權重最高，絕不會被 AI 或更新腳本覆蓋。",
  
  "12": {
    "mandarin": "自訂編號 12 的優化中文...",
    "taiwanese": "自訂編號 12 的道地台語漢字..."
  },
  
  "399-1": {
    "taiwanese": "自訂子項目的台語漢字..."
  },
  
  "principle_1": {
    "title": "自訂核心原則 1 標題",
    "mandarin": "自訂核心原則 1 的中文內容..."
  }
}
```

- **卡牌覆蓋**：以簡報卡牌編號作為 Key（例如 `"12"`、`"399-1"`），亦可使用英文原文作為 Key。
- **原則覆蓋**：以 `"principle_1"` 至 `"principle_20"` 作為 Key。
- **欄位支援**：可僅覆蓋單一欄位（如僅改 `"taiwanese"`），未覆蓋的欄位將自動沿用既有快取庫翻譯。
- **生效方式**：編輯存檔後，執行 `python update_handbook.py` 即可同步產出至 HTML、MD 與 ODT。

---

## 四、網頁四大檢視視角、語音播放與關係圖譜架構

網頁應用程式 (`index.html`) 具備以下核心互動功能與模組：

### 1. 緊湊型雙層導航結構 (Compact 2-Tier Header)
- **結構對稱與空間最佳化**：
  - **第 1 層 (品牌與視角導航)**：左側為精簡品牌標題、副標題與更新時間徽章；右側為「核心原則 (20)」、「簡報卡牌 (458)」、「隨機抽卡」、「關係圖譜 🕸️」四大導航分頁與 `🌓` 深淺色模式切換，徹底消除右上角空白死區。
  - **第 2 層 (欄位開關與即時搜尋)**：左側為「簡報原文 / 優化中文 / 優化英文 / 優化台文 / 對照原則」獨立顯示膠囊與快速模式（全部/純三語/純原文）；右側為即時搜尋輸入框，左右平衡對稱。
  - **垂直高度縮減 60%**：整體導航列高度由 ~240px 大幅降至 ~98px，顯著釋放閱覽視野。

### 2. 優化英文語音朗讀 (Web Speech API)
- 每張簡報卡牌、核心原則卡片、隨機抽卡背面及圖譜詳細資訊抽屜中的「優化英文」旁皆配有專屬發音按鈕（`🔊`）。
- 採用瀏覽器原生 `window.speechSynthesis`，零依賴、純離線、美式標準發音，發音時按鈕切換為 `⏹️` 並伴隨卡片邊框脈衝高亮光暈。

### 3. 雙視角連續自動播放器 (Dual Autoplay Player Bars)
- **原則與卡牌全面支援**：「核心原則 (20)」與「簡報卡牌 (458)」視角頂部均配有專屬的連續自動播放控制器。
- **功能特點**：
  - **一鍵循序朗讀**：一鍵依序朗讀符合條件的原則或卡牌優化英文，項目之間自然停頓 1.0 秒。
  - **視覺平滑置中**：朗讀時自動平滑滾動（Smooth Scroll）將卡片置中，並呈現呼吸發光高亮外框。
  - **智慧篩選聯動**：完全尊重篇章章節、投影片批次與即時搜尋關鍵字，只播放符合篩選的清單。
  - **進度與速率控制**：即時彩條進度百分比與總數比（如 `🎧 朗讀中 原則 1 (1 / 20)`），支援 `0.85x` / `1.0x` / `1.15x` 語速切換，以及上一張、下一張、暫停與停止控制。

### 4. 四種檢視視角
1. **核心原則 (20)**：系統化 20 大原則卡片，包含所屬章節標籤、篇章核心指引金句橫幅、三語文本與關聯對照卡牌跳轉按鈕。
2. **簡報卡牌 (458)**：逐條查閱投影片所有卡牌，支援按投影片批次過濾與即時搜尋。
3. **隨機抽卡**：模擬實體卡牌抽取與 3D 翻牌動畫，背面支援三語對照與個別發音，適合每日職場啟發或自我抽測。
4. **關係圖譜 🕸️ (Connection Graph)**：
   - **零依賴物理引擎**：基於原生 HTML5 Canvas，運用庫倫斥力（推開防重疊）、虎克彈簧力（拉近關聯）、向心重力（防止飄散）與速度阻尼（收斂平衡）進行力導向物理模擬（Force-Directed Simulation）。
   - **拓撲結構 (三層全景架構)**：
     - 6 大篇章 Hub 節點（主題色、半徑 34px）。
     - 20 大核心原則節點（依引用卡牌數動態縮放半徑 18~28px）。
     - 458 張簡報卡牌節點（半徑 6.5px；被多原則共用的「跨界交匯卡」以金色標記並形成跨原則橋接）。
     - 篇章結構虛線 + 簡報卡牌交集重疊實線（線寬依共用數加粗並附計數標籤）+ 跨章節哲學聯動線（Semantic Synergies）+ 原則至卡牌衛星引力連線。
   - **互動控制**：
     - **模式切換**：「🌐 原則骨幹 (20)」與「✨ 全景星系 (458卡牌)」一鍵切換。
     - **衛星展開**：雙擊原則節點或於抽屜點擊「🪐 展開/收合卡牌」，以該原則為中心展開衛星卡牌群。
     - **畫布導航**：支援平移、平滑縮放（0.35x~2.8x）、節點任意拖曳重組、懸停浮動 Tooltip 預覽、懸停高亮路徑、篇章篩選、視角重置與暫停/繼續模擬。
   - **滑出式詳細抽屜 (Sliding Drawer)**：點擊任一原則或卡牌節點即滑出右側抽屜，呈現三語文本、優化英文發音、關聯原則跳轉膠囊與直達卡牌檢視連結。

---

## 五、三語翻譯標準與要求

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

## 六、常見問題排查 (Troubleshooting)

### 1. 新增的項目沒有翻譯或顯示 `[TBD]`
- **原因**：環境中缺少 `GEMINI_API_KEY` 或 API 呼叫受限。
- **解法**：檢查根目錄 `.env` 檔案，確保金鑰有效（系統預設使用 `gemini-3.8-flash`）。或手動將自訂翻譯寫入 `manual_overrides.json` 後重新執行 `python update_handbook.py`。

### 2. 投影片編號跳號或子項目被拆開
- **原因**：簡報在同一張 Slide 中使用了多個不同大小的文字框。
- **解法**：`update_handbook.py` 內建啟發式判斷 `current_num - num > 20` 即判定為子項目並合併。如有特例，可微調 `parse_pptx_items` 中的正則或門檻值。

### 3. GitHub Pages 樣式跑版或載入失敗
- **檢查點**：
  - 確保根目錄存在 `.nojekyll`。
  - 確保 GitHub 倉庫 Settings -> Pages 中的 Source 為 `Deploy from a branch` (Branch: `main` / `/ (root)`).
  - 確保所有靜態資源與字體均使用 HTTPS 外部 CDN 連結。
