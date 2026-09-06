# 工作的管見 (Workplace Principles Handbook) - 專案規範與運作守則

本檔案為 Antigravity 與所有 AI 助理在維護此專案時必須嚴格遵守的核心規則與準則。

---

## 1. 架構核心：單一真實來源 (Single Source of Truth)

- **核心生成器**：[`update_handbook.py`](file:///c:/Users/tu-hs/OneDrive/%E6%96%87%E4%BB%B6/2022_0308_MASA/2022-0708/Projects_antigravity/%E5%B7%A5%E4%BD%9C%E7%9A%84%E7%AE%A1%E8%A6%8B-%E6%9B%B8/update_handbook.py) 是整個工作流的生成中樞。
- **嚴禁單獨修改 HTML**：
  - 切勿僅修改根目錄的 `index.html`，所有 HTML、CSS 與 JavaScript 邏輯**必須同步維護於 `update_handbook.py` 內建的 HTML 模板字串中**。
  - 任何程式碼變更後，務必執行 `python update_handbook.py`，確認回傳代碼為 0，並保證 `index.html`、`工作原則整理與潤稿.md` 及 `工作原則整理與潤稿.odt` 同步產出且毫無遺漏。

---

## 2. 三語對照與翻譯資料庫守則

- **固定三語格式**：
  1. **優化中文 (國語 / Mandarin)**：符合繁體中文商務語境的高品質潤飾。
  2. **優化英文 (English)**：專業、簡潔且道地的職場表達。
  3. **優化台文 (Taiwanese Hokkien)**：必須使用**標準台語漢字**（如「鬥相共」、「省氣力」、「做代誌」、「冗雜」等），蘊含職場處世智慧與民間俗諺感。
- **快取優先原則 (Cache First)**：
  - 翻譯資料一律先比對 `translation_db.json`。
  - 只有在偵測到投影片新增原則或修改英文原文時，才允許呼叫 Gemini API (`gemini-2.5-flash`)。
  - 若無設定 `GEMINI_API_KEY`，系統需自動填入 `[TBD]` 標記，嚴禁中斷工作流或造成程式異常。

---

## 3. 投影片解析與資料處理規範

- **子項目判定啟發法 (Sub-item Heuristic)**：
  - 若投影片中的項目編號突然大幅減小（如 No. 399 後方出現 1, 2, 3），自動判定為上層主項目的子項並合併文字，維持層級整潔。
- **重複編號處理**：
  - 若簡報存在相同編號（例如 Slide 1 的兩個 `17`、`18`），腳本會自動加上 `-1`、`-2` 後綴，確保每張卡牌有唯一的識別 Key。
- **金句 (`☆` 項目) 排除**：
  - 簡報中以 `☆` 包裹的篇章引導金句（例如：`☆利用工具，三觀五覆。☆`）不列入原則卡牌與書籍排版編號，避免干擾 20 條核心原則與 458 條簡報的精準對應。

---

## 4. 網頁卡牌 (`index.html`) 與 RWD 設計標準

- **自包含 (Self-contained)**：
  - 保持單一 HTML 檔案包含所有資料、樣式與腳本，無須 Node.js、npm 或後端伺服器即可離線或線上流暢運作。
- **獨立欄位顯示開關 (Display Toggles)**：
  - 必須支援「簡報原文」、「優化中文」、「優化英文」、「優化台文」與「對照原則」獨立開關。
  - 必須透過 `localStorage` 記住使用者瀏覽偏好。
- **極致 RWD 體驗**：
  - 手機端 (< 768px)：章節篩選器與開關按鈕須維持橫向滑動膠囊（Horizontal Scroll Chips），禁止折行破壞版面。
  - 抽卡模式背面文字若超出高度，必須支援卡片內部平滑滾動（`-webkit-overflow-scrolling: touch`）。

---

## 5. Git 與 GitHub Pages 安全防護

- **機密防護**：
  - 嚴格透過 `.gitignore` 排除 `.env`（內含 API 金鑰）、`__pycache__/`、`*.lnk`、`scratch/` 與暫存企劃檔案。
  - 嚴禁將任何 API 金鑰提交到 Git 倉儲中。
- **GitHub Pages 規範**：
  - 專案根目錄必須維持 `.nojekyll` 檔案。
  - 主分支為 `main`，部署目標為根目錄 `/ (root)`。
  - 遠端倉庫位址：`https://github.com/SinLiongToo/my_priciple_card.git`
  - 公開線上網址：`https://sinliongtoo.github.io/my_priciple_card/`
