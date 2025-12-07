// 檔案名稱：Code.gs

// --- 全域設定 ---
const SPREADSHEET_ID = ''; // 請確認您的 Google Sheet ID

// ==========================================
// 第一區：Web App 控制器 (Controller)
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('複式記帳系統')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAccountsList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Accounts');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift(); 
  return data
    .filter(row => row[1] != '')
    .map(row => ({
      code: row[1],
      name: row[2],
      category: row[3],
      fullName: row[1] + ' - ' + row[2]
    }));
}

// ==========================================
// 第二區：後端核心邏輯 (Backend Logic)
// ==========================================

function saveTransaction(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headerSheet = ss.getSheetByName('JournalEntries');
  const lineSheet = ss.getSheetByName('JournalEntryLines');
  
  if (!headerSheet || !lineSheet) throw new Error('資料庫尚未初始化');

  let totalDebit = 0;
  let totalCredit = 0;
  data.lines.forEach(line => {
    totalDebit += Number(line.debit) || 0;
    totalCredit += Number(line.credit) || 0;
  });
  
  if (Math.abs(totalDebit - totalCredit) > 0.01) {
    throw new Error('借貸不平衡！借方: ' + totalDebit + ', 貸方: ' + totalCredit);
  }

  const entryID = Utilities.getUuid();
  const createdAt = new Date();
  
  headerSheet.appendRow([
    entryID,
    data.date,
    data.description,
    data.reference,
    createdAt,
    'Active' 
  ]);
  
  const lineRows = data.lines.map(line => [
    Utilities.getUuid(),
    entryID,
    line.accountCode,
    Number(line.debit) || 0,
    Number(line.credit) || 0
  ]);
  
  if (lineRows.length > 0) {
    lineSheet.getRange(lineSheet.getLastRow() + 1, 1, lineRows.length, lineRows[0].length).setValues(lineRows);
  }
  
  return { success: true, id: entryID };
}

// ==========================================
// 第三區：系統維護與初始化 (System Setup)
// ==========================================

function initializeDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetsSchema = {
    'Accounts': ['ID', 'Code', 'Name', 'Category', 'Normal_Balance', 'Parent_ID'],
    'FiscalPeriods': ['ID', 'Year', 'Month', 'Start_Date', 'End_Date', 'Status'],
    'JournalEntries': ['ID', 'Date', 'Description', 'Reference_No', 'Created_At', 'Status'],
    'JournalEntryLines': ['ID', 'Journal_Entry_ID', 'Account_Code', 'Debit', 'Credit']
  };

  for (let sheetName in sheetsSchema) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const headers = sheetsSchema[sheetName];
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f3f3');
      sheet.setFrozenRows(1);
    }
  }
}

function seedChartOfAccounts() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Accounts');
  if (!sheet) return;

  const accountsData = [
    ['1100', '1100', '現金及約當現金 (Cash)', 'Asset', 'Debit', ''],
    ['1110', '1110', '銀行存款', 'Asset', 'Debit', '1100'],
    ['1140', '1140', '應收帳款 (A/R)', 'Asset', 'Debit', ''],
    ['1200', '1200', '存貨 (Inventory)', 'Asset', 'Debit', ''],
    ['1250', '1250', '預付費用 (Prepaid)', 'Asset', 'Debit', ''],
    ['1400', '1400', '辦公設備/固定資產', 'Asset', 'Debit', ''],
    ['5100', '5100', '銷貨成本 (COGS)', 'Expense', 'Debit', ''],
    ['6100', '6100', '薪資支出', 'Expense', 'Debit', ''],
    ['6150', '6150', '租金支出', 'Expense', 'Debit', ''],
    ['6200', '6200', '折舊費用', 'Expense', 'Debit', ''],
    ['6250', '6250', '雜項購置/辦公用品', 'Expense', 'Debit', ''],
    ['6280', '6280', '交通費/車資', 'Expense', 'Debit', ''],
    ['6300', '6300', '廣告行銷費', 'Expense', 'Debit', ''],
    ['6350', '6350', '交際費', 'Expense', 'Debit', ''],
    ['6400', '6400', '水電瓦斯費', 'Expense', 'Debit', ''],
    ['2100', '2100', '應付帳款 (A/P)', 'Liability', 'Credit', ''],
    ['2110', '2110', '短期借款', 'Liability', 'Credit', ''],
    ['2140', '2140', '應付費用', 'Liability', 'Credit', ''],
    ['2200', '2200', '預收收入', 'Liability', 'Credit', ''],
    ['2500', '2500', '長期借款', 'Liability', 'Credit', ''],
    ['3100', '3100', '股本/資本', 'Equity', 'Credit', ''],
    ['3300', '3300', '保留盈餘', 'Equity', 'Credit', ''],
    ['3350', '3350', '本期損益', 'Equity', 'Credit', ''],
    ['4100', '4100', '銷貨收入', 'Revenue', 'Credit', ''],
    ['4200', '4200', '服務收入', 'Revenue', 'Credit', ''],
    ['4800', '4800', '利息收入', 'Revenue', 'Credit', '']
  ];

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  sheet.getRange(2, 1, accountsData.length, accountsData[0].length).setValues(accountsData);
}

// ==========================================
// 第四區：報表運算引擎 (Reporting Engine)
// ==========================================

function getTrialBalance() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const headerSheet = ss.getSheetByName('JournalEntries');
  const hData = headerSheet.getDataRange().getValues();
  const validIds = new Set();
  
  for (let i = 1; i < hData.length; i++) {
    const status = hData[i][5]; 
    if (status !== 'Deleted') {
      validIds.add(hData[i][0]);
    }
  }

  const accountsSheet = ss.getSheetByName('Accounts');
  const accountsData = accountsSheet.getDataRange().getValues();
  accountsData.shift(); 
  
  const accountMap = {};
  accountsData.forEach(row => {
    if(row[1]) {
      accountMap[row[1]] = {
        code: row[1], name: row[2], category: row[3], normal: row[4],
        totalDebit: 0, totalCredit: 0
      };
    }
  });

  const lineSheet = ss.getSheetByName('JournalEntryLines');
  if (lineSheet.getLastRow() <= 1) return formatReportData(accountMap);
  
  const linesData = lineSheet.getRange(2, 1, lineSheet.getLastRow() - 1, 5).getValues();

  linesData.forEach(row => {
    const entryId = row[1];
    if (!validIds.has(entryId)) return;

    const code = row[2];
    const debit = Number(row[3]) || 0;
    const credit = Number(row[4]) || 0;
    
    if (accountMap[code]) {
      accountMap[code].totalDebit += debit;
      accountMap[code].totalCredit += credit;
    }
  });

  return formatReportData(accountMap);
}

function formatReportData(accountMap) {
  const report = [];
  for (let code in accountMap) {
    const acc = accountMap[code];
    let balance = (acc.normal === 'Debit') ? (acc.totalDebit - acc.totalCredit) : (acc.totalCredit - acc.totalDebit);
    report.push({
      code: acc.code, name: acc.name, category: acc.category,
      debit: acc.totalDebit, credit: acc.totalCredit, balance: balance
    });
  }
  report.sort((a, b) => String(a.code).localeCompare(String(b.code)));
  return report;
}

function getFinancialReports(year, month) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const lineSheet = ss.getSheetByName('JournalEntryLines');
    const headerSheet = ss.getSheetByName('JournalEntries');
    
    const hData = headerSheet.getDataRange().getValues();
    const dateMap = {};
    const validIds = new Set();

    for(let i=1; i<hData.length; i++) {
      const id = hData[i][0];
      const date = new Date(hData[i][1]);
      const status = hData[i][5]; 
      
      if (status !== 'Deleted') {
        dateMap[id] = date;
        validIds.add(id);
      }
    }

    const lData = lineSheet.getDataRange().getValues();
    const accSheet = ss.getSheetByName('Accounts');
    const accData = accSheet.getDataRange().getValues();
    const accInfo = {};
    for(let i=1; i<accData.length; i++) {
      accInfo[accData[i][1]] = { name: accData[i][2], category: accData[i][3], code: String(accData[i][1]) };
    }

    let periodStart, periodEnd;
    if (month) {
      periodStart = new Date(year, month - 1, 1);
      periodEnd = new Date(year, month, 0, 23, 59, 59);
    } else {
      periodStart = new Date(year, 0, 1);
      periodEnd = new Date(year, 11, 31, 23, 59, 59);
    }

    const income = { revenue: {}, expense: {}, totalRevenue: 0, totalExpense: 0, netIncome: 0 };
    const balance = { assets: {}, liabilities: {}, equity: {}, totalAssets: 0, totalLiabilities: 0, totalEquity: 0 };
    let previousRetainedEarnings = 0; 
    
    const cf = {
      netIncome: 0, depreciation: 0,
      workingCapital: { assetChange: 0, liabilityChange: 0 },
      investing: 0, financing: 0,
      details: { investing: [], financing: [] }
    };

    for(let i=1; i<lData.length; i++) {
      const entryId = lData[i][1];
      if (!validIds.has(entryId)) continue;

      const code = String(lData[i][2]);
      const debit = Number(lData[i][3]) || 0;
      const credit = Number(lData[i][4]) || 0;
      const date = dateMap[entryId];

      if (!date || !accInfo[code]) continue;
      const acc = accInfo[code];
      const netAmount = debit - credit;

      const isPrior = date < periodStart;
      const isCurrent = date >= periodStart && date <= periodEnd;
      const isFuture = date > periodEnd;

      if (isFuture) continue;

      if (acc.category === 'Revenue' || acc.category === 'Expense') {
        if (isPrior) {
          previousRetainedEarnings += (netAmount * -1);
        } else if (isCurrent) {
          if (acc.category === 'Revenue') {
            if (!income.revenue[code]) income.revenue[code] = {name: acc.name, bal: 0};
            income.revenue[code].bal += (credit - debit);
            income.totalRevenue += (credit - debit);
          } else {
            if (!income.expense[code]) income.expense[code] = {name: acc.name, bal: 0};
            income.expense[code].bal += (debit - credit);
            income.totalExpense += (debit - credit);
            if (code === '6200') cf.depreciation += (debit - credit);
          }
        }
      } else {
        if (date <= periodEnd) {
           let list, val;
           if (acc.category === 'Asset') { list = balance.assets; val = netAmount; balance.totalAssets += val; }
           else if (acc.category === 'Liability') { list = balance.liabilities; val = netAmount*-1; balance.totalLiabilities += val; }
           else { list = balance.equity; val = netAmount*-1; balance.totalEquity += val; }
           
           if (!list[code]) list[code] = {name: acc.name, bal: 0};
           list[code].bal += val;
        }

        if (isCurrent && code !== '1100' && code !== '1110') {
           const cashEffect = netAmount * -1; 
           if (acc.category === 'Asset') {
             if ((code.startsWith('14') || code.startsWith('17')) && !acc.name.includes('累計')) {
                cf.investing += cashEffect;
                cf.details.investing.push({name: acc.name, balance: cashEffect});
             } else {
                cf.workingCapital.assetChange += cashEffect;
             }
           } else if (acc.category === 'Liability') {
             if (code.startsWith('25')) {
                cf.financing += cashEffect;
                cf.details.financing.push({name: acc.name, balance: cashEffect});
             } else {
                cf.workingCapital.liabilityChange += cashEffect;
             }
           } else if (acc.category === 'Equity' && code.startsWith('31')) {
                cf.financing += cashEffect;
                cf.details.financing.push({name: acc.name, balance: cashEffect});
           }
        }
      }
    }

    income.netIncome = income.totalRevenue - income.totalExpense;
    if (!balance.equity['3300']) balance.equity['3300'] = {name: '保留盈餘', bal: 0};
    balance.equity['3300'].bal += previousRetainedEarnings;
    balance.totalEquity += previousRetainedEarnings;
    balance.equity['9999'] = {name: '本期淨利 (Net Income)', bal: income.netIncome};
    balance.totalEquity += income.netIncome;

    cf.netIncome = income.netIncome;
    const cashFromOps = cf.netIncome + cf.depreciation + cf.workingCapital.assetChange + cf.workingCapital.liabilityChange;
    const netCashChange = cashFromOps + cf.investing + cf.financing;
    const fmt = (obj) => Object.keys(obj).map(k => ({code:k, name:obj[k].name, balance:obj[k].bal})).sort((a,b)=>a.code.localeCompare(b.code));

    return {
      income: { revenue: fmt(income.revenue), expense: fmt(income.expense), totalRevenue: income.totalRevenue, totalExpense: income.totalExpense, netIncome: income.netIncome },
      balance: { assets: fmt(balance.assets), liabilities: fmt(balance.liabilities), equity: fmt(balance.equity), totalAssets: balance.totalAssets, totalLiabilities: balance.totalLiabilities, totalEquity: balance.totalEquity },
      cashFlow: {
        operating: { netIncome: cf.netIncome, depreciation: cf.depreciation, wcAsset: cf.workingCapital.assetChange, wcLiability: cf.workingCapital.liabilityChange, total: cashFromOps },
        investing: { total: cf.investing, details: cf.details.investing },
        financing: { total: cf.financing, details: cf.details.financing },
        netChange: netCashChange
      },
      equity: { retainedStart: previousRetainedEarnings, netIncome: income.netIncome, capital: balance.equity['3100'] ? balance.equity['3100'].bal : 0, total: balance.totalEquity }
    };
  } catch (e) {
    throw new Error("報表運算錯誤: " + e.message);
  }
}

function exportReportExcel(year, month, mode) {
  const data = getFinancialReports(year, month);
  const tempSS = SpreadsheetApp.create('Temp_Export_Financial_Report_' + new Date().toISOString());
  
  // 建立樣式設定的輔助函式 (避免重複寫四次)
  const formatReportSheet = (sheet, title, headers, dataRows) => {
    // 寫入標題與資料
    const allRows = [headers, ...dataRows];
    const range = sheet.getRange(1, 1, allRows.length, headers.length);
    range.setValues(allRows);
    
    // 1. 標題列樣式
    sheet.getRange(1, 1, 1, headers.length)
         .setBackground('#1565c0') // 較亮的藍色
         .setFontColor('#ffffff')
         .setFontWeight('bold')
         .setHorizontalAlignment('center');

    // 2. 凍結第一列
    sheet.setFrozenRows(1);
    
    // 3. 金額欄位格式 (最後一欄通常是金額)
    sheet.getRange(2, headers.length, dataRows.length, 1).setNumberFormat('#,##0');
    
    // 4. 自動調整欄寬
    sheet.autoResizeColumns(1, headers.length);
    
    // 5. 針對「小計」、「總計」、「本期淨利」等關鍵字加粗背景
    for (let i = 0; i < dataRows.length; i++) {
      const firstCellVal = String(dataRows[i][0]);
      // 如果第一欄包含特定關鍵字，將整行加粗並上淡灰色背景
      if (firstCellVal.includes('小計') || firstCellVal.includes('總計') || firstCellVal.includes('淨利') || firstCellVal.includes('---')) {
        const rowRange = sheet.getRange(i + 2, 1, 1, headers.length); // i+2 因為有標題列且 i 從 0 開始
        rowRange.setFontWeight('bold');
        if(!firstCellVal.includes('---')) {
           rowRange.setBackground('#f5f5f5');
        }
      }
    }
  };

  // A. 損益表
  const sheetInc = tempSS.getSheets()[0];
  sheetInc.setName('綜合損益表');
  let rowsInc = [];
  rowsInc.push(['--- 收入 ---', '', '']);
  data.income.revenue.forEach(i => rowsInc.push([i.code, i.name, i.balance]));
  rowsInc.push(['小計', '收入總計', data.income.totalRevenue]);
  rowsInc.push(['--- 費用 ---', '', '']);
  data.income.expense.forEach(i => rowsInc.push([i.code, i.name, i.balance]));
  rowsInc.push(['小計', '費用總計', data.income.totalExpense]);
  rowsInc.push(['', '本期淨利', data.income.netIncome]);
  formatReportSheet(sheetInc, '綜合損益表', ['科目代碼', '科目名稱', '金額'], rowsInc);

  // B. 資產負債表
  const sheetBal = tempSS.insertSheet('資產負債表');
  let rowsBal = [];
  rowsBal.push(['--- 資產 ---', '', '']);
  data.balance.assets.forEach(i => rowsBal.push([i.code, i.name, i.balance]));
  rowsBal.push(['總計', '資產總計', data.balance.totalAssets]);
  rowsBal.push(['--- 負債 ---', '', '']);
  data.balance.liabilities.forEach(i => rowsBal.push([i.code, i.name, i.balance]));
  rowsBal.push(['總計', '負債總計', data.balance.totalLiabilities]);
  rowsBal.push(['--- 權益 ---', '', '']);
  data.balance.equity.forEach(i => rowsBal.push([i.code, i.name, i.balance]));
  rowsBal.push(['總計', '權益總計', data.balance.totalEquity]);
  formatReportSheet(sheetBal, '資產負債表', ['科目代碼', '科目名稱', '金額'], rowsBal);

  // C. 現金流量表
  const sheetCF = tempSS.insertSheet('現金流量表');
  let rowsCF = [];
  rowsCF.push(['營業活動', '']);
  rowsCF.push(['本期淨利', data.cashFlow.operating.netIncome]);
  rowsCF.push(['加：折舊', data.cashFlow.operating.depreciation]);
  rowsCF.push(['營運資產變動', data.cashFlow.operating.wcAsset]);
  rowsCF.push(['營運負債變動', data.cashFlow.operating.wcLiability]);
  rowsCF.push(['小計 (營業活動淨現金流)', data.cashFlow.operating.total]);
  rowsCF.push(['投資活動', '']);
  data.cashFlow.investing.details.forEach(d => rowsCF.push([d.name, d.balance]));
  rowsCF.push(['小計 (投資活動淨現金流)', data.cashFlow.investing.total]);
  rowsCF.push(['籌資活動', '']);
  data.cashFlow.financing.details.forEach(d => rowsCF.push([d.name, d.balance]));
  rowsCF.push(['小計 (籌資活動淨現金流)', data.cashFlow.financing.total]);
  rowsCF.push(['總計 (本期現金淨變動)', data.cashFlow.netChange]);
  formatReportSheet(sheetCF, '現金流量表', ['項目', '金額'], rowsCF);
  
  // D. 權益變動表
  const sheetEq = tempSS.insertSheet('權益變動表');
  let rowsEq = [];
  rowsEq.push(['期初保留盈餘', data.equity.retainedStart]);
  rowsEq.push(['加：本期淨利', data.equity.netIncome]);
  rowsEq.push(['加：股本餘額', data.equity.capital]);
  rowsEq.push(['期末權益總額', data.equity.total]);
  formatReportSheet(sheetEq, '權益變動表', ['項目', '金額'], rowsEq);

  return getExportUrl_(tempSS.getId());
}

function exportInvoiceExcel(startDateStr, endDateStr) {
  // 取得資料
  const invoices = getInvoicesByDateRange(startDateStr, endDateStr);
  
  // 建立暫存 Spreadsheet
  const tempSS = SpreadsheetApp.create('Temp_Export_Invoices_' + new Date().toISOString());
  const sheet = tempSS.getSheets()[0];
  sheet.setName('發票明細');

  // 取得科目名稱對照
  const acctList = getAccountsList();
  const acctMap = {};
  acctList.forEach(a => acctMap[a.code] = a.name);

  // 定義標題
  const headers = ['系統單號 (System ID)', '日期', '單號 (Reference)', '摘要/說明', '科目代碼', '科目名稱', '借方金額', '貸方金額'];
  const rows = [headers];

  // 整理資料
  invoices.forEach(inv => {
    inv.lines.forEach(line => {
      rows.push([
        inv.id,
        inv.date,
        inv.reference || '',
        inv.description,
        line.accountCode,
        acctMap[line.accountCode] || '',
        line.debit,
        line.credit
      ]);
    });
  });

  // 寫入資料
  if (rows.length > 0) {
    const range = sheet.getRange(1, 1, rows.length, rows[0].length);
    range.setValues(rows);
    
    // --- 開始排版美化 ---
    
    // 1. 標題列樣式 (深藍底白字、置中、粗體)
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#2c3e50')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setHorizontalAlignment('center');
    
    // 2. 凍結第一列
    sheet.setFrozenRows(1);

    // 3. 設定資料欄位格式
    // B欄 (日期): yyyy-mm-dd
    sheet.getRange(2, 2, rows.length - 1, 1).setNumberFormat('yyyy-mm-dd');
    // G欄~H欄 (金額): 千分位數值
    sheet.getRange(2, 7, rows.length - 1, 2).setNumberFormat('#,##0');

    // 4. 畫框線 (細黑線)
    range.setBorder(true, true, true, true, true, true, '#dcdcdc', SpreadsheetApp.BorderStyle.SOLID);

    // 5. 自動調整欄寬
    sheet.autoResizeColumns(1, headers.length);
  }

  return getExportUrl_(tempSS.getId());
}
// ==========================================
// 第六區：發票管理與 CRUD 核心 (Invoice Management)
// ==========================================

function getInvoicesByDateRange(startDateStr, endDateStr) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headerSheet = ss.getSheetByName('JournalEntries');
  const lineSheet = ss.getSheetByName('JournalEntryLines');
  
  const lastRow = headerSheet.getLastRow();
  if (lastRow <= 1) return [];

  const headers = headerSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  
  const start = new Date(startDateStr); start.setHours(0, 0, 0, 0);
  const end = new Date(endDateStr); end.setHours(23, 59, 59, 999);

  const filteredHeaders = headers.filter(row => {
    const rowDate = new Date(row[1]);
    const status = row[5];
    return rowDate >= start && rowDate <= end && status !== 'Deleted';
  });

  if (filteredHeaders.length === 0) return [];

  const lineLastRow = lineSheet.getLastRow();
  let allLines = [];
  if (lineLastRow > 1) {
    allLines = lineSheet.getRange(2, 1, lineLastRow - 1, 5).getValues();
  }

  filteredHeaders.sort((a, b) => new Date(b[1]) - new Date(a[1]));

  return filteredHeaders.map(h => {
    const id = h[0];
    const myLines = allLines.filter(l => l[1] === id).map(l => ({
      accountCode: l[2], debit: Number(l[3]), credit: Number(l[4])
    }));
    const total = myLines.reduce((sum, l) => sum + l.debit, 0);
    return {
      id: id,
      date: Utilities.formatDate(new Date(h[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      description: h[2], reference: h[3], total: total, lines: myLines
    };
  });
}

function getTrashedInvoices() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headerSheet = ss.getSheetByName('JournalEntries');
  const lineSheet = ss.getSheetByName('JournalEntryLines');
  
  const lastRow = headerSheet.getLastRow();
  if (lastRow <= 1) return [];

  const headers = headerSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  
  const filteredHeaders = headers.filter(row => row[5] === 'Deleted');
  
  if (filteredHeaders.length === 0) return [];

  const lineLastRow = lineSheet.getLastRow();
  let allLines = [];
  if (lineLastRow > 1) allLines = lineSheet.getRange(2, 1, lineLastRow - 1, 5).getValues();

  filteredHeaders.sort((a, b) => new Date(b[1]) - new Date(a[1]));

  return filteredHeaders.map(h => {
    const id = h[0];
    const myLines = allLines.filter(l => l[1] === id).map(l => ({
      accountCode: l[2], debit: Number(l[3]), credit: Number(l[4])
    }));
    const total = myLines.reduce((sum, l) => sum + l.debit, 0);
    return {
      id: id,
      date: Utilities.formatDate(new Date(h[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      description: h[2], reference: h[3], total: total, lines: myLines
    };
  });
}

function deleteInvoice(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headerSheet = ss.getSheetByName('JournalEntries');
  const data = headerSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      headerSheet.getRange(i + 1, 6).setValue('Deleted');
      return { success: true };
    }
  }
  throw new Error('找不到該筆資料');
}

function restoreInvoice(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headerSheet = ss.getSheetByName('JournalEntries');
  const data = headerSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      headerSheet.getRange(i + 1, 6).setValue('Active');
      return { success: true };
    }
  }
  throw new Error('找不到該筆資料');
}

function hardDeleteInvoice(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headerSheet = ss.getSheetByName('JournalEntries');
  const lineSheet = ss.getSheetByName('JournalEntryLines');

  const lineData = lineSheet.getDataRange().getValues();
  for (let i = lineData.length - 1; i >= 1; i--) {
    if (lineData[i][1] === id) {
      lineSheet.deleteRow(i + 1);
    }
  }

  const headerData = headerSheet.getDataRange().getValues();
  for (let i = headerData.length - 1; i >= 1; i--) {
    if (headerData[i][0] === id) {
      headerSheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  throw new Error('找不到該筆資料');
}

function updateInvoice(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headerSheet = ss.getSheetByName('JournalEntries');
  const lineSheet = ss.getSheetByName('JournalEntryLines');
  const id = data.id;

  const headerData = headerSheet.getDataRange().getValues();
  let headerRowIndex = -1;
  for (let i = 1; i < headerData.length; i++) {
    if (headerData[i][0] === id) {
      headerRowIndex = i + 1;
      break;
    }
  }

  if (headerRowIndex === -1) throw new Error('找不到該筆發票 ID');

  headerSheet.getRange(headerRowIndex, 2).setValue(data.date);
  headerSheet.getRange(headerRowIndex, 3).setValue(data.description);
  headerSheet.getRange(headerRowIndex, 4).setValue(data.reference);

  const lineData = lineSheet.getDataRange().getValues();
  for (let i = lineData.length - 1; i >= 1; i--) {
    if (lineData[i][1] === id) {
      lineSheet.deleteRow(i + 1);
    }
  }

  const newLines = data.lines.map(line => [
    Utilities.getUuid(), id, line.accountCode,
    Number(line.debit) || 0, Number(line.credit) || 0
  ]);

  if (newLines.length > 0) {
    lineSheet.getRange(lineSheet.getLastRow() + 1, 1, newLines.length, newLines[0].length).setValues(newLines);
  }

  return { success: true };
}

function reverseTransaction(originalId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headerSheet = ss.getSheetByName('JournalEntries');
  const lineSheet = ss.getSheetByName('JournalEntryLines');
  
  const headerData = headerSheet.getDataRange().getValues();
  let originalHeader = null;
  for (let i = 1; i < headerData.length; i++) {
    if (headerData[i][0] == originalId) {
      originalHeader = headerData[i];
      break;
    }
  }
  if (!originalHeader) throw new Error('找不到原始傳票 ID');

  const lineData = lineSheet.getDataRange().getValues();
  const originalLines = lineData.filter(row => row[1] == originalId);
  if (originalLines.length === 0) throw new Error('找不到原始傳票明細');

  const newEntryID = Utilities.getUuid();
  const today = new Date(); 
  
  headerSheet.appendRow([
    newEntryID, today, '[作廢] ' + originalHeader[2], originalHeader[3], new Date(), 'Active'
  ]);

  const newLines = originalLines.map(line => [
    Utilities.getUuid(), newEntryID, line[2], line[4], line[3]
  ]);

  if (newLines.length > 0) {
    lineSheet.getRange(lineSheet.getLastRow() + 1, 1, newLines.length, newLines[0].length).setValues(newLines);
  }
  return { success: true, newId: newEntryID };
}

function runYearEndClosing(closingDate) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headerSheet = ss.getSheetByName('JournalEntries');
  const lineSheet = ss.getSheetByName('JournalEntryLines');
  
  const trialData = getTrialBalance();
  
  const closingLines = [];
  let totalDebit = 0;
  let totalCredit = 0;
  
  trialData.forEach(acc => {
    if ((acc.category === 'Revenue' || acc.category === 'Expense') && Math.abs(acc.balance) > 0.01) {
      if (acc.balance > 0) {
        if (acc.normal === 'Debit') {
          closingLines.push({ accountCode: acc.code, debit: 0, credit: acc.balance });
          totalCredit += acc.balance;
        } else {
          closingLines.push({ accountCode: acc.code, debit: acc.balance, credit: 0 });
          totalDebit += acc.balance;
        }
      }
    }
  });
  
  if (closingLines.length === 0) throw new Error('沒有需要結帳的科目餘額');

  const retainedEarningsCode = '3300'; 
  const diff = totalDebit - totalCredit;
  
  if (Math.abs(diff) > 0.01) {
    if (diff > 0) {
      closingLines.push({ accountCode: retainedEarningsCode, debit: 0, credit: diff });
    } else {
      closingLines.push({ accountCode: retainedEarningsCode, debit: Math.abs(diff), credit: 0 });
    }
  }

  const entryID = Utilities.getUuid();
  headerSheet.appendRow([
    entryID, closingDate, '年度結帳分錄 (Closing Entry)', 'CLOSE-YEAR', new Date(), 'Active'
  ]);
  
  const newLinesRows = closingLines.map(line => [
    Utilities.getUuid(), entryID, line.accountCode, line.debit, line.credit
  ]);
  
  lineSheet.getRange(lineSheet.getLastRow() + 1, 1, newLinesRows.length, newLinesRows[0].length).setValues(newLinesRows);
  
  return { success: true, count: closingLines.length };
}

function clearAllTransactions() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headerSheet = ss.getSheetByName('JournalEntries');
  const lineSheet = ss.getSheetByName('JournalEntryLines');
  
  if (headerSheet.getLastRow() > 1) {
    headerSheet.getRange(2, 1, headerSheet.getLastRow() - 1, headerSheet.getLastColumn()).clearContent();
  }
  
  if (lineSheet.getLastRow() > 1) {
    lineSheet.getRange(2, 1, lineSheet.getLastRow() - 1, lineSheet.getLastColumn()).clearContent();
  }
  
  return '系統已重置！所有交易紀錄已清空。';
}
// 輔助：產生下載連結
function getExportUrl_(fileId) {
  // export?format=xlsx 是 Google Drive 的標準轉換參數
  return 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
}
