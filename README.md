# Google Apps Script 複式記帳系統 (Double-Entry Bookkeeping System)

這是一個基於 **Google Apps Script (GAS)** 與 **Google Sheets** 構建的輕量級雲端會計系統。它遵循嚴格的複式記帳原則 (Double-entry bookkeeping)，適合中小型專案、個人工作室或學習會計系統開發使用。

## 🌟 功能特色 (Features)

*   **雲端化與免安裝**：完全運行於 Google 生態系，無需架設伺服器。
*   **複式記帳核心**：
    *   借貸平衡檢查 (Debit = Credit)。
    *   支援多行分錄 (Split Transactions)。
*   **完整的 CRUD 管理**：
    *   新增、查詢、修改、刪除傳票。
    *   **垃圾桶機制**：支援軟刪除 (Soft Delete)，可隨時還原誤刪資料。
*   **自動化財務報表**：
    *   綜合損益表 (Income Statement)
    *   資產負債表 (Balance Sheet)
    *   現金流量表 (Cash Flow Statement - 間接法)
    *   權益變動表 (Statement of Changes in Equity)
    *   *支援年度與自訂月份區間查詢營運成果。*
*   **Excel 匯出**：一鍵將發票明細或財務報表匯出為格式精美的 Excel 檔案。
*   **現代化 UI**：響應式設計 (RWD)，支援手機與電腦操作。

## 🛠️ 安裝與部署教學 (Setup Guide)

### 1. 建立 Google Sheet
1.  建立一個新的 Google Sheet。
2.  記下網址中的 **Spreadsheet ID** (例如：`docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit` 中的 `YOUR_SPREADSHEET_ID`)。

### 2. 部署程式碼
1.  開啟 Google Sheet 的 **擴充功能 (Extensions)** > **Apps Script**。
2.  將 `code.gs` 的內容複製並貼上到 Apps Script 編輯器中的 `Code.gs`。
3.  將 `Index.html` 的內容複製並貼上，建立一個同名的 HTML 檔案。
4.  **重要**：在 `code.gs` 第 4 行，填入您的 Spreadsheet ID：
    ```javascript
    const SPREADSHEET_ID = '您的_SPREADSHEET_ID_貼在這裡';
    ```

### 3. 初始化系統 (自動建立工作表與標題)
1.  **這是全自動的！您不需要手動建立工作表或輸入標題。**
2.  在 Apps Script 編輯器上方工具列，選擇 `initializeDatabase` 函式並點擊 **執行**。
    *   執行後，系統會自動在您的 Google Sheet 中建立以下工作表自並填入標題列：
        *   `Accounts`
        *   `JournalEntries`
        *   `JournalEntryLines`
        *   `FiscalPeriods`
3.  接著選擇 `seedChartOfAccounts` 函式並點擊 **執行**。
    *   這將會自動寫入預設的 **會計科目表 (Chart of Accounts)** (如：現金、銷貨收入、薪資支出等)。

### 4. 發布 Web App
1.  點擊右上角的 **部署 (Deploy)** > **新增部署 (New deployment)**。
2.  選擇類型為 **網頁應用程式 (Web app)**。
3.  設定如下：
    *   **執行身分 (Execute as)**: `Me` (即您的帳號)。
    *   **誰可以存取 (Who has access)**: `Only myself` (僅自己) 或根據需求設定。
4.  點擊 **部署**，您將獲得一個專屬的 Web App 網址。

## 📂 專案結構 (Project Structure)

| 檔案 | 說明 |
| :--- | :--- |
| **code.gs** | 後端核心。包含路由控制器、資料庫邏輯、報表運算引擎及 Excel 匯出功能。 |
| **Index.html** | 前端介面。單頁應用程式 (SPA)，包含發票登錄、查詢、報表檢視與垃圾桶管理介面。 |
| **MockDataGenerator.gs** | (選用) 用於生成測試資料，方便開發者快速驗證報表功能。 |

## 📊 資料庫架構 (Schema)

系統會自動建立以下資料表結構，若您需要手動維護資料，請參考以下標題設定：

### 1. Accounts (會計科目表)
存放所有可用的會計科目。
*   **標題列 (Row 1):** `ID`, `Code`, `Name`, `Category`, `Normal_Balance`, `Parent_ID`
*   **範例:** `1100`, `1100`, `現金 (Cash)`, `Asset`, `Debit`, ``

### 2. JournalEntries (傳票主檔)
存放每一筆交易的日期、摘要與狀態。
*   **標題列 (Row 1):** `ID`, `Date`, `Description`, `Reference_No`, `Created_At`, `Status`
*   **說明:** `Status` 欄位用於控制軟刪除 (Active/Deleted)。

### 3. JournalEntryLines (傳票明細檔)
存放每一筆交易的借貸方金額。
*   **標題列 (Row 1):** `ID`, `Journal_Entry_ID`, `Account_Code`, `Debit`, `Credit`
*   **說明:** 透過 `Journal_Entry_ID` 與主檔關聯。

### 4. FiscalPeriods (會計期間 - 保留用)
用於未來擴充關帳與期間管理功能。
*   **標題列 (Row 1):** `ID`, `Year`, `Month`, `Start_Date`, `End_Date`, `Status`

## ⚠️ 免責聲明 (Disclaimer)

本系統僅供學術研究、個人記帳或輔助使用。作者不對因使用本軟體造成的任何財務損失或數據錯誤負責。在用於正式商業用途前，請務必諮詢專業會計師。

---
Developed by [ycchou]
