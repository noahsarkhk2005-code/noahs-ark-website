/**
 * Tally 虛擬訂單測試
 * 模擬 Tally 寫入 NA_Tickets / NA_Merch（不經真實表單）
 *
 * 執行：
 *   createTallyVirtualOrders     — 新增一組虛擬購票 + 商品單
 *   createTallyVirtualTickets    — 只加購票測試
 *   createTallyVirtualMerch      — 只加商品測試
 *   clearTallyVirtualOrders      — 刪除 Submission ID 以 TEST- 開頭的列
 *
 * 之後可跑：syncTallyTicketsToOrders（若已安裝）
 */

var VT_ = {
  TICKETS_SHEET: 'NA_Tickets',
  MERCH_SHEET: 'NA_Merch',
  // 假付款截圖（公開 placeholder，僅測試用）
  FAKE_PROOF: 'https://via.placeholder.com/600x400.png?text=TEST+Payment+Proof',
  EMAIL: 'test-virtual@noahsark.local',
  PHONE: '85290000000',
  METAL_PASS: '' // 測試通常留空，避免會員衝突
};

function rangeInclusive_(sh, r1, c1, r2, c2) {
  return sh.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1);
}

function createTallyVirtualOrders() {
  var t = createTallyVirtualTickets_();
  var m = createTallyVirtualMerch_();
  SpreadsheetApp.getUi().alert(
    '虛擬訂單已建立\n\n' +
    '【購票 NA_Tickets】\n' + t + '\n\n' +
    '【商品 NA_Merch】\n' + m + '\n\n' +
    '下一步：\n' +
    '1) 檢查 NA_Tickets / NA_Merch 新列\n' +
    '2) 執行 syncTallyTicketsToOrders（若有）\n' +
    '3) AppSheet Sync 看待審核\n' +
    '4) 測完可執行 clearTallyVirtualOrders'
  );
}

function createTallyVirtualTickets() {
  var msg = createTallyVirtualTickets_();
  SpreadsheetApp.getUi().alert('虛擬購票單：\n' + msg);
}

function createTallyVirtualMerch() {
  var msg = createTallyVirtualMerch_();
  SpreadsheetApp.getUi().alert('虛擬商品單：\n' + msg);
}

function createTallyVirtualTickets_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(VT_.TICKETS_SHEET);
  if (!sh) throw new Error('找不到分頁 ' + VT_.TICKETS_SHEET);

  ensureTicketHeaders_(sh);

  var stamp = Utilities.formatDate(new Date(), 'Asia/Hong_Kong', 'MMddHHmmss');
  var now = new Date();

  // 三張虛擬票單：早鳥 / 預售 / 會員特級
  var cases = [
    {
      sid: 'TEST-T-EB-' + stamp,
      name: '虛擬早鳥測試',
      type: '🎫 早鳥門票 HKD300',
      qty: 2,
      price: 300
    },
    {
      sid: 'TEST-T-ADV-' + stamp,
      name: '虛擬預售測試',
      type: '🎫 預售 HKD350',
      qty: 1,
      price: 350
    },
    {
      sid: 'TEST-T-MEM-' + stamp,
      name: '虛擬會員票測試',
      type: '會員特級優惠 HKD300',
      qty: 1,
      price: 300
    }
  ];

  var lines = [];
  var i;
  for (i = 0; i < cases.length; i++) {
    var c = cases[i];
    var total = c.qty * c.price;
    // A–L：Submission ID … 門票類別 / 數量 / 單價 / 截圖
    var row = [
      c.sid,                    // A Submission ID
      'RESP-TEST',              // B Respondent ID
      now,                      // C Submitted at
      total,                    // D Total Amount
      c.name,                   // E Name
      VT_.EMAIL,                // F Email
      VT_.PHONE,                // G Tel
      VT_.METAL_PASS,           // H Metal Pass
      c.type,                   // I 門票類別
      c.qty,                    // J 數量
      c.price,                  // K 單價
      VT_.FAKE_PROOF            // L Payment Capture
    ];
    sh.appendRow(row);
    lines.push(c.sid + ' · ' + c.type + ' ×' + c.qty + ' = HK$' + total);
  }
  return lines.join('\n');
}

function createTallyVirtualMerch_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(VT_.MERCH_SHEET);
  if (!sh) throw new Error('找不到分頁 ' + VT_.MERCH_SHEET);

  ensureMerchHeaders_(sh);
  var lastCol = Math.max(sh.getLastColumn(), 15);

  var stamp = Utilities.formatDate(new Date(), 'Asia/Hong_Kong', 'MMddHHmmss');
  var now = new Date();

  var cases = [
    { sid: 'TEST-M-TEE-' + stamp, name: '虛擬 Tee 測試', cat: '服飾', qty: 1, price: 280 },
    { sid: 'TEST-M-TOWEL-' + stamp, name: '虛擬毛巾測試', cat: '毛巾', qty: 2, price: 80 },
    { sid: 'TEST-M-PACK-' + stamp, name: '虛擬 Pack 測試', cat: '套裝 Pack', qty: 1, price: 200 }
  ];

  var lines = [];
  var i;
  for (i = 0; i < cases.length; i++) {
    var c = cases[i];
    var total = c.qty * c.price;
    // 對齊常見 NA_Merch 寬度：A–O + 可能的空欄
    var row = new Array(lastCol);
    var j;
    for (j = 0; j < lastCol; j++) row[j] = '';

    row[0] = c.sid;
    row[1] = 'RESP-TEST';
    row[2] = now;
    row[3] = total;
    row[4] = c.name;
    row[5] = VT_.EMAIL;
    row[6] = VT_.PHONE;
    row[7] = VT_.METAL_PASS;
    row[8] = c.cat;           // I 商品分欄
    row[9] = c.qty;           // J 數量
    row[10] = c.price;        // K 單價
    // 付款截圖：找表頭或預設 O (15)
    var proofCol = findHeaderCol_(sh, ['請上傳付款截圖', 'Payment Capture', '付款截圖']);
    if (proofCol < 0) proofCol = 14; // O = index 14
    if (proofCol < lastCol) row[proofCol] = VT_.FAKE_PROOF;

    sh.appendRow(row);
    lines.push(c.sid + ' · ' + c.cat + ' ×' + c.qty + ' = HK$' + total);
  }
  return lines.join('\n');
}

/** 刪除 TEST- 開頭的虛擬列 */
function clearTallyVirtualOrders() {
  var ss = SpreadsheetApp.getActive();
  var n1 = clearTestRows_(ss.getSheetByName(VT_.TICKETS_SHEET));
  var n2 = clearTestRows_(ss.getSheetByName(VT_.MERCH_SHEET));
  // 同步 log
  var log = ss.getSheetByName('_TallyTicketSyncLog');
  var n3 = 0;
  if (log) n3 = clearTestRows_(log);

  SpreadsheetApp.getUi().alert(
    '已清除虛擬訂單\nNA_Tickets: ' + n1 + ' 列\nNA_Merch: ' + n2 + ' 列\nSyncLog: ' + n3 + ' 列'
  );
}

function clearTestRows_(sh) {
  if (!sh || sh.getLastRow() < 2) return 0;
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  var data = rangeInclusive_(sh, 2, 1, lastRow, Math.max(lastCol, 1)).getValues();
  var removed = 0;
  // 由下往上刪
  var r;
  for (r = data.length - 1; r >= 0; r--) {
    var sid = String(data[r][0] || '').trim();
    if (sid.indexOf('TEST-') === 0) {
      sh.deleteRow(r + 2);
      removed++;
    }
  }
  return removed;
}

function ensureTicketHeaders_(sh) {
  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
    sh.clear();
    sh.appendRow([
      'Submission ID', 'Respondent ID', 'Submitted at', 'Total Amount',
      '💀 Name/稱呼', '📩 Email/電郵', '☎️ Tel/電話號碼', 'Metal Pass 會員編號',
      '門票類別', '數量', '單價', 'Payment Capture / 付款截圖'
    ]);
    return;
  }
  // 若 I 還不是「門票類別」，只補 A1
  var a1 = String(sh.getRange(1, 1).getValue() || '').trim();
  if (!a1) sh.getRange(1, 1).setValue('Submission ID');
  var i1 = String(sh.getRange(1, 9).getValue() || '').trim();
  if (!i1 || i1.indexOf('門票') < 0 && i1.indexOf('早鳥') < 0 && i1.indexOf('類別') < 0) {
    // 若尚未跑過 setupNaIjkDropdowns，仍寫入新表頭（不覆蓋已有下拉邏輯時可手改）
    if (!i1) {
      sh.getRange(1, 9).setValue('門票類別');
      sh.getRange(1, 10).setValue('數量');
      sh.getRange(1, 11).setValue('單價');
    }
  }
}

function ensureMerchHeaders_(sh) {
  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
    sh.clear();
    sh.appendRow([
      'Submission ID', 'Respondent ID', 'Submitted at', 'Total Amount',
      '💀 Name/稱呼', '📩 Email/電郵', '☎️ Tel/電話號碼', 'Metal Pass 會員編號',
      '商品分欄', '數量', '單價',
      'Ark BigPack HKD200', 'Ark TinyPack HKD60', 'Ark Tote Bag HKD120',
      '請上傳付款截圖 / Please upload your payment screenshot'
    ]);
    return;
  }
  var a1 = String(sh.getRange(1, 1).getValue() || '').trim();
  if (!a1) sh.getRange(1, 1).setValue('Submission ID');
}

function findHeaderCol_(sh, names) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return -1;
  var headers = rangeInclusive_(sh, 1, 1, 1, lastCol).getValues()[0];
  var c, n;
  for (c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '');
    for (n = 0; n < names.length; n++) {
      if (h.indexOf(names[n]) >= 0) return c;
    }
  }
  return -1;
}
