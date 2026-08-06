/**
 * NA_Tickets / NA_Merch：I·J·K 改為下拉分類（欄位數量不變）
 *
 * NA_Tickets（12 欄不變）
 *   I 門票類別  ← 下拉（會員特級優惠 / 早鳥 / 預售）
 *   J 數量
 *   K 單價（選類別後可手填或由腳本帶入）
 *
 * NA_Merch（欄位數不變）
 *   I 商品分欄  ← 下拉（服飾 / 配件 / 毛巾 / Pack / 袋類 或 商品名）
 *   J 數量
 *   K 單價
 *
 * 用法：Apps Script 貼上 → 執行 setupNaIjkDropdowns
 * 會嘗試把舊「各品項數量」欄的資料遷移到新 I/J/K
 */

var TICKET_TYPES_ = [
  '會員特級優惠 HKD300',
  '🎫 早鳥門票 HKD300',
  '🎫 預售 HKD350'
];

var TICKET_PRICES_ = {
  '會員特級優惠 HKD300': 300,
  '🎫 早鳥門票 HKD300': 300,
  '🎫 預售 HKD350': 350
};

// 商品分欄（細類）— 對齊 Order_Categories
var MERCH_CATEGORIES_ = [
  '服飾',
  '配件',
  '毛巾',
  '套裝 Pack',
  '袋類'
];

// 舊 merch I/J/K 品名 → 分欄
var MERCH_OLD_TO_CAT_ = {
  'Ark T-Shirt HKD280': '服飾',
  'Ark Tower/毛巾 HKD80': '毛巾',
  'Ark Keychain/鎖匙扣 HKD120': '配件',
  'Ark BigPack HKD200': '套裝 Pack',
  'Ark TinyPack HKD60': '套裝 Pack',
  'Ark Tote Bag HKD120': '袋類'
};

var MERCH_PRICES_ = {
  '服飾': 280,
  '配件': 120,
  '毛巾': 80,
  '套裝 Pack': 200,
  '袋類': 120
};

function setupNaIjkDropdowns() {
  var ss = SpreadsheetApp.getActive();
  var report = [];

  report.push(setupTicketsIjk_(ss));
  report.push(setupMerchIjk_(ss));

  SpreadsheetApp.getUi().alert(report.join('\n\n'));
}

function setupTicketsIjk_(ss) {
  var sh = ss.getSheetByName('NA_Tickets');
  if (!sh) return 'NA_Tickets: 找不到分頁';

  var lastCol = Math.max(sh.getLastColumn(), 12);
  var lastRow = Math.max(sh.getLastRow(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  // 確保至少 12 欄
  if (lastCol < 12) {
    sh.getRange(1, lastCol + 1, 1, 12 - lastCol).setValues([
      new Array(12 - lastCol).fill('')
    ]);
    lastCol = 12;
    headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  }

  // 記住舊 IJK 表頭（遷移用）
  var oldI = String(headers[8] || '').trim();
  var oldJ = String(headers[9] || '').trim();
  var oldK = String(headers[10] || '').trim();

  // 新表頭（欄位數不變：仍是 I J K）
  sh.getRange(1, 9).setValue('門票類別');   // I
  sh.getRange(1, 10).setValue('數量');      // J
  sh.getRange(1, 11).setValue('單價');      // K

  // 若 A1 空白，補 Submission ID
  if (!String(headers[0] || '').trim()) {
    sh.getRange(1, 1).setValue('Submission ID');
  }

  // 遷移資料列
  var migrated = 0;
  if (lastRow >= 2) {
    var data = sh.getRange(2, 1, lastRow, lastCol).getValues();
    var outI = [];
    var outJ = [];
    var outK = [];
    var r;
    for (r = 0; r < data.length; r++) {
      var qtyI = toNum_(data[r][8]);
      var qtyJ = toNum_(data[r][9]);
      var qtyK = toNum_(data[r][10]);
      var type = '';
      var qty = 0;
      var price = '';

      // 已是新格式？I 已是文字類別
      var iVal = data[r][8];
      if (iVal !== '' && iVal != null && isNaN(Number(iVal)) && String(iVal).indexOf('http') < 0) {
        type = String(iVal);
        qty = toNum_(data[r][9]) || 1;
        price = toNum_(data[r][10]) || (TICKET_PRICES_[type] || '');
      } else {
        // 舊：三欄數量 → 取第一個有數量的票種
        if (qtyI > 0) {
          type = oldI || TICKET_TYPES_[0];
          qty = qtyI;
        } else if (qtyJ > 0) {
          type = oldJ || TICKET_TYPES_[1];
          qty = qtyJ;
        } else if (qtyK > 0) {
          type = oldK || TICKET_TYPES_[2];
          qty = qtyK;
        }
        if (type) {
          // 對齊標準名稱
          type = matchTicketType_(type);
          price = TICKET_PRICES_[type] || '';
          migrated++;
        }
      }

      outI.push([type || '']);
      outJ.push([qty || '']);
      outK.push([price || '']);
    }
    sh.getRange(2, 9, lastRow, 9).setValues(outI);
    sh.getRange(2, 10, lastRow, 10).setValues(outJ);
    sh.getRange(2, 11, lastRow, 11).setValues(outK);
  }

  // 下拉：門票類別（I 欄，預留 500 列）
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(TICKET_TYPES_, true)
    .setAllowInvalid(true)
    .setHelpText('選擇門票類別')
    .build();
  sh.getRange(2, 9, 500, 9).setDataValidation(rule);

  // 表頭樣式
  sh.getRange(1, 9, 1, 3).setFontWeight('bold').setBackground('#fce8e6');

  return 'NA_Tickets: I=門票類別 J=數量 K=單價（下拉已設，遷移 ' + migrated + ' 列）';
}

function setupMerchIjk_(ss) {
  var sh = ss.getSheetByName('NA_Merch');
  if (!sh) return 'NA_Merch: 找不到分頁';

  var lastCol = Math.max(sh.getLastColumn(), 15);
  var lastRow = Math.max(sh.getLastRow(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  var oldI = String(headers[8] || '').trim();
  var oldJ = String(headers[9] || '').trim();
  var oldK = String(headers[10] || '').trim();

  // 新表頭 — 欄位數不變（L 以後原樣保留）
  sh.getRange(1, 9).setValue('商品分欄');  // I
  sh.getRange(1, 10).setValue('數量');     // J
  sh.getRange(1, 11).setValue('單價');     // K

  var migrated = 0;
  if (lastRow >= 2) {
    var data = sh.getRange(2, 1, lastRow, lastCol).getValues();
    var outI = [];
    var outJ = [];
    var outK = [];
    var r;
    for (r = 0; r < data.length; r++) {
      var qtyI = toNum_(data[r][8]);
      var qtyJ = toNum_(data[r][9]);
      var qtyK = toNum_(data[r][10]);
      var cat = '';
      var qty = 0;
      var price = '';

      var iVal = data[r][8];
      if (iVal !== '' && iVal != null && isNaN(Number(iVal)) && String(iVal).indexOf('http') < 0) {
        cat = String(iVal);
        qty = toNum_(data[r][9]) || 1;
        price = toNum_(data[r][10]) || (MERCH_PRICES_[cat] || '');
      } else {
        if (qtyI > 0) {
          cat = MERCH_OLD_TO_CAT_[oldI] || '服飾';
          qty = qtyI;
          price = guessMerchPrice_(oldI);
        } else if (qtyJ > 0) {
          cat = MERCH_OLD_TO_CAT_[oldJ] || '毛巾';
          qty = qtyJ;
          price = guessMerchPrice_(oldJ);
        } else if (qtyK > 0) {
          cat = MERCH_OLD_TO_CAT_[oldK] || '配件';
          qty = qtyK;
          price = guessMerchPrice_(oldK);
        }
        // 若 IJK 都空，但 L/M/N 有量，不搬（保持 L+ 原欄）
        if (cat) migrated++;
      }

      outI.push([cat || '']);
      outJ.push([qty || '']);
      outK.push([price || '']);
    }
    sh.getRange(2, 9, lastRow, 9).setValues(outI);
    sh.getRange(2, 10, lastRow, 10).setValues(outJ);
    sh.getRange(2, 11, lastRow, 11).setValues(outK);
  }

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(MERCH_CATEGORIES_, true)
    .setAllowInvalid(true)
    .setHelpText('選擇商品分欄')
    .build();
  sh.getRange(2, 9, 500, 9).setDataValidation(rule);

  sh.getRange(1, 9, 1, 3).setFontWeight('bold').setBackground('#e6f4ea');

  return 'NA_Merch: I=商品分欄 J=數量 K=單價（下拉已設，遷移 ' + migrated + ' 列；L 起原商品欄保留）';
}

function matchTicketType_(raw) {
  var s = String(raw || '');
  var i;
  for (i = 0; i < TICKET_TYPES_.length; i++) {
    if (s === TICKET_TYPES_[i]) return TICKET_TYPES_[i];
  }
  if (s.indexOf('會員') >= 0 || s.indexOf('Metal') >= 0) return TICKET_TYPES_[0];
  if (s.indexOf('早鳥') >= 0 || s.indexOf('Early') >= 0) return TICKET_TYPES_[1];
  if (s.indexOf('預售') >= 0 || s.indexOf('Advanced') >= 0) return TICKET_TYPES_[2];
  return s;
}

function guessMerchPrice_(header) {
  var m = String(header || '').match(/HKD?\s*(\d+)/i);
  return m ? Number(m[1]) : '';
}

function toNum_(v) {
  if (v === '' || v == null) return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var n = Number(String(v).replace(/[^0-9.\-]/g, ''));
  return isFinite(n) ? n : 0;
}

/** 選單 */
function onOpen_NaDropdowns() {
  SpreadsheetApp.getUi()
    .createMenu('NOAHSARK-欄位')
    .addItem('設定 NA IJK 下拉（門票類別／商品分欄）', 'setupNaIjkDropdowns')
    .addToUi();
}
