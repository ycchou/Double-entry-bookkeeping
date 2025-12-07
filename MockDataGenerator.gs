// 檔案名稱：MockDataGenerator.gs

/**
 * 生成 300 筆擬真的測試資料
 * 包含：銷貨收入、費用支出、採購、薪資等情境
 * 執行完畢後請重新整理網頁查看報表
 */
function generateRealWorldMockData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headerSheet = ss.getSheetByName('JournalEntries');
  const lineSheet = ss.getSheetByName('JournalEntryLines');

  // --- 設定資料庫 ---
  // 為了讓資料更真實，我們定義一些隨機的客戶、供應商和摘要
  const customers = ['台積電', '聯發科', '中華電信', 'Google Taiwan', '蝦皮購物', '王小明', '陳小姐', '未來科技公司'];
  const suppliers = ['PChome 企業採購', 'AWS 雲端服務', '台灣大車隊', '房東林先生', 'Costco 好市多', 'Adobe Creative Cloud', 'Facebook 廣告'];
  const expenseTypes = [
    { name: '辦公室租金', code: '6150', amount: 30000, supplier: '房東林先生' },
    { name: '員工薪資', code: '6100', amount: [40000, 150000], supplier: '銀行轉帳' }, // 範圍
    { name: 'Facebook 廣告費', code: '6300', amount: [500, 5000], supplier: 'Meta Platforms' },
    { name: 'Google 關鍵字廣告', code: '6300', amount: [1000, 8000], supplier: 'Google Ads' },
    { name: '辦公文具用品', code: '6250', amount: [200, 2000], supplier: 'PChome 24h' },
    { name: '計程車資', code: '6250', amount: [150, 800], supplier: 'Uber/Taxi' },
    { name: '電費', code: '6250', amount: [2000, 6000], supplier: '台灣電力公司' }
  ];

  const headers = [];
  const lines = [];
  
  // 設定日期範圍 (今年 1/1 ~ 今天)
  const endDate = new Date();
  const startDate = new Date(endDate.getFullYear(), 0, 1);

  // --- 1. 先建立一筆「期初資金」確保現金流為正 ---
  const initId = Utilities.getUuid();
  headers.push([initId, startDate, '股東現金增資 (期初開帳)', 'INIT-001', new Date()]);
  lines.push([Utilities.getUuid(), initId, '1100', 500000, 0]); // 借：現金 50萬
  lines.push([Utilities.getUuid(), initId, '3100', 0, 500000]); // 貸：股本

  // --- 2. 迴圈生成 299 筆隨機交易 ---
  for (let i = 0; i < 299; i++) {
    const date = getRandomDate(startDate, endDate);
    const id = Utilities.getUuid();
    const type = Math.random(); // 決定交易類型
    
    let description = '';
    let ref = '';
    let debitCode = '';
    let creditCode = '';
    let amount = 0;

    // 情境 A: 銷貨收入 (50% 機率)
    if (type < 0.5) {
      const customer = customers[Math.floor(Math.random() * customers.length)];
      amount = Math.floor(Math.random() * 20000) + 1000; // 1000 ~ 21000
      description = `銷售服務給 ${customer}`;
      ref = 'INV-' + Math.floor(Math.random() * 10000);
      
      // 隨機決定是收現(1100) 還是 賒帳(1140)
      if (Math.random() > 0.3) {
        debitCode = '1100'; // 收現
        description += ' (現金)';
      } else {
        debitCode = '1140'; // 應收
        description += ' (月結)';
      }
      creditCode = '4100'; // 銷貨收入
    }
    // 情境 B: 支付費用 (40% 機率)
    else if (type < 0.9) {
      const exp = expenseTypes[Math.floor(Math.random() * expenseTypes.length)];
      if (Array.isArray(exp.amount)) {
        amount = Math.floor(Math.random() * (exp.amount[1] - exp.amount[0])) + exp.amount[0];
      } else {
        amount = exp.amount;
      }
      description = `支付${exp.name} - ${exp.supplier}`;
      ref = 'EXP-' + Math.floor(Math.random() * 10000);
      debitCode = exp.code;
      
      // 隨機決定付現(1100) 還是 應付(2100)
      if (Math.random() > 0.2) {
        creditCode = '1100'; // 現金支付
      } else {
        creditCode = '2100'; // 應付帳款
      }
    }
    // 情境 C: 採購資產/設備 (10% 機率)
    else {
      amount = Math.floor(Math.random() * 30000) + 5000;
      description = '採購電腦設備與硬體';
      ref = 'ASSET-' + Math.floor(Math.random() * 1000);
      debitCode = '1400'; // PP&E
      creditCode = '1100'; // 現金
    }

    // 存入陣列 (主檔)
    headers.push([
      id,
      date,
      description,
      ref,
      new Date()
    ]);

    // 存入陣列 (明細 - 借方)
    lines.push([
      Utilities.getUuid(),
      id,
      debitCode,
      amount,
      0
    ]);

    // 存入陣列 (明細 - 貸方)
    lines.push([
      Utilities.getUuid(),
      id,
      creditCode,
      0,
      amount
    ]);
  }

  // --- 3. 批次寫入資料庫 (Sort by Date) ---
  // 為了好看，我們先依照日期排序一下
  headers.sort((a, b) => new Date(a[1]) - new Date(b[1]));

  // 寫入 Header
  if (headers.length > 0) {
    headerSheet.getRange(headerSheet.getLastRow() + 1, 1, headers.length, headers[0].length).setValues(headers);
  }

  // 寫入 Lines
  // 注意：Lines 不用特別排，因為是用 Journal_Entry_ID 關聯的
  if (lines.length > 0) {
    lineSheet.getRange(lineSheet.getLastRow() + 1, 1, lines.length, lines[0].length).setValues(lines);
  }

  Logger.log('成功生成 300 筆擬真交易資料！');
}

/**
 * 輔助函式：取得範圍內的隨機日期
 */
function getRandomDate(start, end) {
  return new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
}
