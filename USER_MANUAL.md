# 使用手冊 (USER MANUAL)

本手冊提供「工作的管見」工作流與網頁的詳細使用說明。

---

## 國語說明 (Mandarin Guide)

### 第一部分：Python 工作流操作
當您未來修改或更新了 `工作原則-書.pptx.pptx` 投影片後，請依照以下步驟更新手冊與網頁：

1. **設定 API 金鑰（選用）**：
   - 如果您只修改了投影片的格式或順序，或沒有新增項目，則不需要金鑰。
   - 如果您**新增了投影片原則**或**修改了英文原文**，系統需要呼叫 API 進行翻譯。請在專案根目錄下建立一個名為 `.env` 的檔案，寫入您的金鑰：
     ```env
     GEMINI_API_KEY=您的金鑰內容
     ```
2. **執行更新腳本**：
   - 開啟您的終端機（PowerShell 或 CMD），切換至本專案目錄。
   - 執行指令：
     ```bash
     python update_handbook.py
     ```
   - 腳本會解析投影片、比對翻譯庫。若有新修改的項目，會自動呼叫 Gemini 進行三語翻譯，寫入快取，並重新產生手冊與網頁。

### 第二部分：卡牌網頁操作
雙擊開啟 `index.html` 網頁後，您可以進行以下操作：
1. **切換檢視視角**：
   - **核心原則 (20)**：瀏覽整理歸納後的 20 條工作大原則，適合系統性閱讀。
   - **簡報卡牌 (452)**：展開全部 452 條簡報原始項目與潤稿卡片。
   - **隨機抽卡**：卡牌翻牌遊戲。點擊「隨機抽一張卡」會隨機抽取一條原則，卡片正面為簡報英文原文，點擊卡片即可 3D 翻轉至背面查看三語翻譯。
2. **即時搜尋與篩選**：
   - 在「核心原則」與「簡報卡牌」頁面中，可以在右上角搜尋框輸入編號、中文、英文或台語關鍵字進行即時過濾。
   - 可點擊上方的過濾標籤（如「Slide 1」、「第一篇」），快速篩選特定範圍的卡牌。
3. **獨立欄位顯示切換（開關）**：
   - 在頂部「顯示切換」欄中，您可以獨立開啟或關閉：
     - **簡報原文**：投影片原本書寫的英文原則筆記。
     - **優化中文**：商業潤稿中文。
     - **優化英文**：精煉職場英文。
     - **優化台文**：台語漢字智慧。
     - **對照原則**：核心原則 1~20 的徽章標籤。
   - **快速情境**：點擊「全部顯示」、「純三語」或「純原文」快速套用配置。
   - 您的設定會自動儲存於瀏覽器，下次進入時無需重複設定。
4. **主題切換**：
   - 點擊右上角的 🌓 按鈕，可在護眼的深色模式與明亮的淺色模式間自由切換。

### 第三部分：發布至 GitHub Pages (公開網站)
1. **本地提交 Git**：
   在專案目錄下執行：
   ```bash
   git init
   git add .
   git commit -m "feat: 工作的管見卡牌網頁與 GitHub Pages 支援"
   ```
2. **在 GitHub 建立 Repository 並推送**：
   ```bash
   git branch -M main
   git remote add origin https://github.com/SinLiongToo/my_priciple_card.git
   git push -u origin main
   ```
3. **在 GitHub 開啟 Pages**：
   進入倉庫的 **Settings** -> **Pages**，將 **Source** 設為 **Deploy from a branch**（Branch: `main` / `/ (root)`）後按 Save。
   1~2 分鐘後即可使用專屬網址隨時隨地用手機或電腦閱讀卡牌：
   `https://sinliongtoo.github.io/my_priciple_card/`

---

## English Guide

### Part 1: Python Workflow Operation
If you modify or update the `工作原則-書.pptx.pptx` PowerPoint file in the future, follow these steps to update the handbook and web app:

1. **Set up API Key (Optional)**:
   - If you only reordered or adjusted existing slides without editing the English texts, no API key is required.
   - If you **added new principles** or **modified the original English texts**, the system needs to call the API for translation. Create a `.env` file in the project root and add your key:
     ```env
     GEMINI_API_KEY=your_gemini_api_key_here
     ```
2. **Run the Update Script**:
   - Open your terminal (PowerShell or CMD) and navigate to the project directory.
   - Run the command:
     ```bash
     python update_handbook.py
     ```
   - The script will parse the slides, match them against the database, translate new items using Gemini, save them to the cache, and regenerate both the Markdown handbook and the HTML page.

### Part 2: Card Web App Navigation
Double-click `index.html` to open it in a browser, then enjoy these interactive features:
1. **Switch Views**:
   - **Core Principles (20)**: Browse the 20 synthesized principles categorized by chapters.
   - **Slide Cards (452)**: Expand all 452 raw slide entries with polished translations.
   - **Random Card**: A card-drawing interface. Click 'Draw Card' to draw a random principle. The card front shows the original text; click the card to flip it with a 3D animation and reveal the trilingual translations.
2. **Search and Filter**:
   - In the 'Core Principles' and 'Slide Cards' views, use the search input to filter cards instantly by number, English, Mandarin, or Taiwanese keywords.
   - Click filter tags (e.g., 'Slide 1', 'Chapter 1') to focus on specific sections.
3. **Field Visibility Toggles**:
   - In the "Display Toggles" bar at the top, independently show or hide:
     - **Original**: The raw principle note written on the slide.
     - **Polished Mandarin**: Natural Traditional Chinese business translation.
     - **Polished English**: Professional refined English version.
     - **Taiwanese**: Idiomatic Taiwanese Hokkien wisdom.
     - **Principles**: Core principle badge tags and links.
   - **Presets**: Click "All", "Trilingual Only", or "Original Only" for quick switching.
   - Preferences are automatically saved in `localStorage`.
4. **Theme Switch**:
   - Click the 🌓 icon in the header to toggle between Sleek Dark Mode and Elegant Light Mode.

### Part 3: Deploying to GitHub Pages (Live Website)
1. **Commit locally with Git**:
   ```bash
   git init
   git add .
   git commit -m "feat: trilingual card app with GitHub Pages setup"
   ```
2. **Create repository on GitHub & push**:
   ```bash
   git branch -M main
   git remote add origin https://github.com/SinLiongToo/my_priciple_card.git
   git push -u origin main
   ```
3. **Enable GitHub Pages**:
   Go to **Settings** -> **Pages** in your GitHub repository, choose **Deploy from a branch** (Branch: `main` / `/ (root)`), then click Save.
   Your live site will be ready at: `https://sinliongtoo.github.io/my_priciple_card/`.

---

## 台語漢字說明 (Taiwanese Guide)

### 第一部分：Python 工作流操作
假使您以後有修改抑是更新 `工作原則-書.pptx.pptx` 投影片，請照下跤的步驟來更新手冊同網頁：

1. **設定 API 金鑰（無一定要）**：
   - 假使您只有調整投影片的順序，抑是無加新的內容，就毋免金鑰。
   - 假使您有**加新的投影片原則**或**改英文原文**，系統需要呼叫 API 翻譯。請佇專案目錄底下一本建立 `.env` 檔案，寫入金鑰：
     ```env
     GEMINI_API_KEY=您的金鑰內容
     ```
2. **執行更新指令碼**：
   - 開啟終端機（PowerShell 或 CMD），切換到專案目錄。
   - 執行指令：
     ```bash
     python update_handbook.py
     ```
   - 程式會解析投影片、比對快取資料庫。若有新改的項目，會自動呼叫 Gemini 進行三語翻譯、寫入快取，並重新產生手冊與網頁。

### 第二部分：卡牌網頁操作
點兩下開啟 `index.html` 網頁後，您可以按呢操作：
1. **切換瀏覽角度**：
   - **核心原則 (20)**：瀏覽整理好的 20 條工作大原則，適合系統性閱讀。
   - **簡報卡牌 (452)**：展開全部 452 條簡報原始項目佮潤稿卡片。
   - **隨機抽卡**：翻牌遊戲。點擊「隨機抽一張卡」會隨機抽一條原則，卡片正面是英文原文，點卡片即可 3D 翻面看三語對照。
2. **搜尋與過濾**：
   - 佇「核心原則」與「簡報卡牌」頁面，佇右上角搜尋框輸入編號、中、英、台語關鍵字進行即時搜尋。
   - 點擊上方的過濾標籤（像「Slide 1」、「第一篇」），會快速篩選特定範圍的卡牌。
3. **獨立欄位顯示開關（自由開關）**：
   - 佇頂頭「顯示切換」欄，會使分別開關：
     - **簡報原文**：投影片原本寫的英文手記。
     - **優化中文**：商業潤稿國語。
     - **優化英文**：道地的英文。
     - **優化台文**：親切接地氣的台語漢字。
     - **對照原則**：核心原則徽章佮連結。
   - **快速模式**：點「全部顯示」、「純三語」抑是「純原文」會快速套用。
   - 瀏覽器會自動記持您的設定，後擺入來毋免重設。
4. **主題切換**：
   - 點右上角 🌓 按鈕，會佇深色模式同淺色模式中自由切換，保護目睭。

### 第三部分：發布到 GitHub Pages 做公開網站
1. **佇本地 Git 提交**：
   ```bash
   git init
   git add .
   git commit -m "feat: 工作的管見卡牌網頁與 GitHub Pages 支援"
   ```
2. **佇 GitHub 開新倉庫並推送**：
   ```bash
   git branch -M main
   git remote add origin https://github.com/SinLiongToo/my_priciple_card.git
   git push -u origin main
   ```
3. **佇 GitHub 開啟 Pages 功能**：
   入去倉庫的 **Settings** -> **Pages**，將 **Source** 選做 **Deploy from a branch**（Branch: `main` / `/ (root)`）並點 Save。
   一兩分鐘了後，就會有個人網址，用手機抑是電腦攏會用得隨時抽卡閱讀：
   `https://sinliongtoo.github.io/my_priciple_card/`
