/**
 * NA_Tickets / NA_Merch：I·J·K 改為下拉分類（欄位數量不變）
 *
 * 注意：Sheet.getRange(row, column, numRows, numColumns)
 * 第三、四參數是「列數、欄數」，不是結束列/結束欄。
 *
 * 用法：貼上 → 執行 setupNaIjkDropdowns
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

var MERCH_CATEGORIES_ = [
  '服飾',
  '配件',
  '毛巾',
  '套裝 Pack',
  '袋類'
];

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

/** 含首尾的範圍：r1,c1 到 r2,c2（1-based） */
function rangeInclusive_(sh, r1, c1, r2, c2) {
  var numRows = r2 - r1 + 1;
  var numCols = c2 - c1 + 1;
  if (numRows < 1 || numCols < 1) {
    throw new Error('rangeInclusive_ 無效: ' + r1 + ',' + c1 + ' → ' + r2 + ',' + c2);
  }
  return sh.getRange(r1, c1, numRows, numCols);
}

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

  // 補足至少 12 欄
  if (sh.getLastColumn() < 12) {
    var need = 12 - sh.getLastColumn();
    sh.insertColumnsAfter(sh.getLastColumn() || 1, need);
    lastCol = 12;
  }

  var headers = rangeInclusive_(sh, 1, 1, 1, lastCol).getValues()[0];
  var oldI = String(headers[8] || '').trim();
  var oldJ = String(headers[9] || '').trim();
  var oldK = String(headers[10] || '').trim();

  // 新表頭 I J K
  sh.getRange(1, 9).setValue('門票類別');
  sh.getRange(1, 10).setValue('數量');
  sh.getRange(1, 11).setValue('單價');

  if (!String(headers[0] || '').trim()) {
    sh.getRange(1, 1).setValue('Submission ID');
  }

  var migrated = 0;
  if (lastRow >= 2) {
    var numDataRows = lastRow - 1;
    var data = rangeInclusive_(sh, 2, 1, lastRow, lastCol).getValues();
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

      var iVal = data[r][8];
      if (iVal !== '' && iVal != null && isNaN(Number(iVal)) && String(iVal).indexOf('http') < 0) {
        type = String(iVal);
        qty = toNum_(data[r][9]) || 1;
        price = toNum_(data[r][10]) || (TICKET_PRICES_[type] || '');
      } else {
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
          type = matchTicketType_(type);
          price = TICKET_PRICES_[type] || '';
          migrated++;
        }
      }

      outI.push([type || '']);
      outJ.push([qty || '']);
      outK.push([price || '']);
    }

    // 單欄寫入：numRows = data 列數，numColumns = 1
    rangeInclusive_(sh, 2, 9, lastRow, 9).setValues(outI);
    rangeInclusive_(sh, 2, 10, lastRow, 10).setValues(outJ);
    rangeInclusive_(sh, 2, 11, lastRow, 11).setValues(outK);
  }

  // 下拉：I 欄 500 列
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(TICKET_TYPES_, true)
    .setAllowInvalid(true)
    .setHelpText('選擇門票類別')
    .build();
  rangeInclusive_(sh, 2, 9, 501, 9).setDataValidation(rule);

  rangeInclusive_(sh, 1, 9, 1, 11).setFontWeight('bold').setBackground('#fce8e6');

  return 'NA_Tickets: I=門票類別 J=數量 K=單價（下拉已設，遷移 ' + migrated + ' 列）';
}

function setupMerchIjk_(ss) {
  var sh = ss.getSheetByName('NA_Merch');
  if (!sh) return 'NA_Merch: 找不到分頁';

  var lastCol = Math.max(sh.getLastColumn(), 15);
  var lastRow = Math.max(sh.getLastRow(), 1);
  var headers = rangeInclusive_(sh, 1, 1, 1, lastCol).getValues()[0];

  var oldI = String(headers[8] || '').trim();
  var oldJ = String(headers[9] || '').trim();
  var oldK = String(headers[10] || '').trim();

  sh.getRange(1, 9).setValue('商品分欄');
  sh.getRange(1, 10).setValue('數量');
  sh.getRange(1, 11).setValue('單價');

  var migrated = 0;
  if (lastRow >= 2) {
    var data = rangeInclusive_(sh, 2, 1, lastRow, lastCol).getValues();
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
        if (cat) migrated++;
      }

      outI.push([cat || '']);
      outJ.push([qty || '']);
      outK.push([price || '']);
    }

    rangeInclusive_(sh, 2, 9, lastRow, 9).setValues(outI);
    rangeInclusive_(sh, 2, 10, lastRow, 10).setValues(outJ);
    rangeInclusive_(sh, 2, 11, lastRow, 11).setValues(outK);
  }

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(MERCH_CATEGORIES_, true)
    .setAllowInvalid(true)
    .setHelpText('選擇商品分欄')
    .build();
  rangeInclusive_(sh, 2, 9, 501, 9).setDataValidation(rule);

  rangeInclusive_(sh, 1, 9, 1, 11).setFontWeight('bold').setBackground('#e6f4ea');

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
