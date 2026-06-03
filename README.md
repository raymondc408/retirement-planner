# 提早退休規劃工具 (Early Retirement Planner)

**互動網頁版（推薦給別人試用）**：直接在瀏覽器使用，無需安裝。支援即時編輯年度開支表、多回報率比較、Simple Model 走勢表、起始資本敏感度分析、完整 Excel 匯出、中英雙語切換。

**本地 Excel 產生器**：也保留了舊版，可批次產生報告。

核心特色：完全參數化、單一年度開支表驅動所有計算、可隨時修改任何年份金額、精準重現你圖片中的 Simple Model 格式。

## 目前功能 (v0.1)

- 輸入：生活開支、兩個子女教育費、按揭剩餘年期
- 支援多種真實報酬率情境比較（3%、3.5%、4%、4.5%、5%）
- 自動計算「50歲退休當下需要多少本金」
- 完整年度資產走勢模擬（考慮通脹 + 子女階段性教育費 + 按揭）
- 自動產生多張圖表與格式化表格
- 清楚列出所有假設，方便審計

## 如何使用

1. 安裝依賴（只需一次）：
   ```powershell
   pip install openpyxl pandas numpy
   ```

2. **推薦做法**：第一次執行後會自動產生 `parameters.xlsx`
   - 打開 `parameters.xlsx`
   - 在**黃色單元格**修改你的數字：
     - 生活開支、按揭每月供款、總按揭剩餘年期
     - 子女教育費排程（表格形式，很容易加減階段）
     - 要比較的被動收入報酬率（列表）
     - 要產生哪幾個退休年齡版本（例如 5, 7, 9 年後）
   - 儲存檔案

3. 執行：
   ```powershell
   python generate_retirement_excel.py
   ```

之後你**只需要改 parameters.xlsx**，再跑一次 Python 就會產生更新後的報告。非常方便隨時測試不同情境。

---

## Streamlit 互動版 (推薦！)

如果你想要**即時修改、滑桿 + 即時表格編輯**的體驗（不用一直存檔重跑）：

**重要**：因為你用的是 Microsoft Store 版本的 Python，`streamlit` 指令通常不會直接在 PATH 裡。

正確啟動方式（在 retirement_planner 資料夾下）：

```powershell
cd C:\Users\RAYMOND\retirement_planner
python -m streamlit run streamlit_app.py
```

這行會用正確的 python 來啟動 streamlit。

瀏覽器打開後：
- 左側 sidebar 設定年齡、退休年份列表（例如 2027,2031）、回報率等。
- 主要區域的「年度開支表」可以用滑鼠直接編輯任何一年的 Kate 學費、Damon 學費、生活開支、按揭。
- 下面會即時顯示「資本走勢表 (Simple Model)」，格式非常接近你上傳的那張圖（Capital / Passive Income / Expense / Gain/Loss）。
- 改完表格或參數後點「重新計算」或等它反應即可看到新結果。
- 右上角有下載目前情境成 Excel 的按鈕。

3. 在網頁上：
   - 直接在「年度開支表」用滑鼠點擊編輯任何年份的 Kate 學費、Damon 學費、生活開支、按揭。
   - 左側改退休年份列表、回報率等。
   - 下面會即時顯示「資本走勢表 (Simple Model)」——格式很像你上傳的那張圖 (Capital / Passive Income / Expense / Gain/Loss)。
   - 有圖表 + 多退休年份所需本金比較。
   - 右上角可以下載目前情境的 Excel。

這個版本體驗最好，隨時改隨時看結果。Excel 產生器還是保留給你想存檔多個版本的時候用。

---

## 線上給別人試用（部署到網頁）

如果你想讓其他人**不用安裝任何東西**就能直接在瀏覽器試用這個工具，有幾種非常簡單的方式：

### 推薦方式 1：Streamlit Community Cloud（最簡單、官方、免費）

1. 把 `retirement_planner` 這個資料夾（包含 `streamlit_app.py`、`requirements.txt`、`README.md`）整個 push 到 GitHub，建立一個 **公開 (public)** 的 repository。

2. 去 [https://share.streamlit.io](https://share.streamlit.io) 或 [Streamlit Community Cloud](https://streamlit.io/cloud) 用 GitHub 登入。

3. **強烈建議**：點 **New app** 後，從列表中**直接選擇你的 GitHub repo**（不要手動貼 URL），這樣系統會自動偵測分支和檔案路徑，比較不會出錯。

4. 如果系統要你手動填（像你現在看到的畫面）：
   - **Repository**：貼上你的完整 repo 網址，例如 `https://github.com/你的帳號/retirement-planner`
   - **Branch**：**一定要改成 `main`**（現在 GitHub 預設新 repo 都是 `main`，不是 `master`）
   - **Main file path**：`streamlit_app.py` （注意：把 retirement_planner 資料夾的內容推上去後，檔案會在 GitHub repo 最上層，所以 Main file path 填 `streamlit_app.py`）
   - 先去你的 GitHub 網頁打開 repo，看實際資料夾結構再決定。

5. 點 Deploy，等 1-3 分鐘就會有公開網址。

**重要**：一定要先把程式碼成功 push 到 GitHub，分支也要是 `main`，檔案路徑要完全正確，否則會出現「This file does not exist」或「This branch does not exist」的錯誤。

以後只要你 push 新 code 到 GitHub，雲端就會自動更新。

優點：
- 完全免費（公開使用）
- 不用管伺服器
- 支援直接下載 Excel（功能完全正常）
- 雙語切換也正常

### 推薦方式 2：Hugging Face Spaces

1. 去 [https://huggingface.co/spaces](https://huggingface.co/spaces) 登入（可用 GitHub）。
2. 建立新 Space，選擇 **Streamlit** 作為 SDK。
3. 用 GitHub repo 連接，或直接上傳檔案。
4. 把 `requirements.txt` 放在根目錄或 Space 裡即可。

Hugging Face 的介面比較漂亮，適合做公開 demo。

### 其他選擇
- **Render.com**：免費額度充足，部署也很快。
- **Railway** 或 **Fly.io**：更進階一點。

### 部署前建議準備

- 已經幫你加了 `requirements.txt`（包含 streamlit, pandas, openpyxl, plotly 等）。
- 建議在 GitHub repo 根目錄加上 `.gitignore`（我已幫你建立，避免上傳 __pycache__、output/ 資料夾）。
- 可以把 `parameters.xlsx` 和 `output/` 資料夾的檔案先刪掉或忽略再 commit。
- 在 app 裡面可以加一行小提示，例如在最下面放 GitHub 連結，讓別人知道可以自己 clone 改。

### 本地測試部署設定（可選）

在 `retirement_planner` 資料夾新增一個 `.streamlit/config.toml` 可以設定主題、port 等，但對雲端部署不是必要。

想讓別人試用最快就是用上面 Streamlit Cloud 或 Hugging Face，幾分鐘就能上線。

---

## 如何在自己電腦繼續使用（本地版）

## 重要假設與提醒（務必閱讀）

- 這是**簡化模型**，僅供參考，不是專業財務意見。
- 目前假設報酬率為「每年固定真實報酬」（已扣通脹後的實質增長）。實際投資有波動、序列風險。
- 生活費與教育費每年按通脹調整。
- 沒有考慮稅務、醫療突發、樓價變動、遺產等。
- 子女教育費時間表需你根據實際情況調整（程式內有清楚參數）。
- 強烈建議多跑幾次不同假設，觀察「所需本金」對報酬率與通脹的敏感度。

## 後續計劃

- 加入累積期模擬（現在到 50 歲每月可儲多少）
- Monte Carlo 隨機模擬（顯示成功率區間）
- 轉成 Streamlit 互動網頁版（你之後想要的話我可以繼續做）

## 聯絡 / 修改建議

有任何想調整的參數、想加的功能，或發現模型問題，隨時告訴我，我可以快速修改。

---

**免責聲明**：本工具產生的數字僅供個人參考，實際退休規劃請諮詢持牌財務顧問。投資有風險，過往表現不代表未來結果。