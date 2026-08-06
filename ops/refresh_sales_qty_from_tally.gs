/**
 * Order_Categories 由 Tally 訂單彙總：
 *   K 欄 = 銷售數量
 *   L 欄 = 單件貨品總收入  (= 銷售數量 × list_price)
 *
 * 資料來源：NA_Tickets / NA_Merch
 * 執行：refreshSalesQtyFromTally
 */

var SQ_ = {
  CAT_SHEET: 'Order_Categories',
  TICKETS: 'NA_Tickets',
  MERCH: 'NA_Merch',
  QTY_COL_NAME: '銷售數量',
  REVENUE_COL_NAME: '單件貨品總收入',
  /** K = 11, L = 12（1-based） */
  QTY_COL_1BASED: 11,
  REVENUE_COL_1BASED: 12,
  /** J = list_price 預設第 10 欄 */
  PRICE_COL_1BASED: 10
};

function rangeInc_(sh, r1, c1, r2, c2) {
  return sh.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1);
}

function refreshSalesQtyFromTally() {
  var ss = SpreadsheetApp.getActive();
  var cat = ss.getSheetByName(SQ_.CAT_SHEET);
  if (!cat) throw new Error('找不到 Order_Categories');

  var needCols = SQ_.REVENUE_COL_1BASED;
  var lastCol = Math.max(cat.getLastColumn(), needCols);
  var lastRow = cat.getLastRow();
  if (lastRow < 2) throw new Error('Order_Categories 沒有資料');

  if (cat.getLastColumn() < needCols) {
    cat.insertColumnsAfter(cat.getLastColumn() || 1, needCols - cat.getLastColumn());
    lastCol = needCols;
  }

  var headers = rangeInc_(cat, 1, 1, 1, lastCol).getValues()[0];
  var idx = headerIndexMap_(headers);

  var qtyCol1 = SQ_.QTY_COL_1BASED;       // K
  var revCol1 = SQ_.REVENUE_COL_1BASED;   // L
  var priceCol0 = idx.list_price != null ? idx.list_price : (SQ_.PRICE_COL_1BASED - 1);
  var levelCol0 = idx.level != null ? idx.level : 2;

  // 表頭 K / L
  cat.getRange(1, qtyCol1).setValue(SQ_.QTY_COL_NAME);
  cat.getRange(1, qtyCol1).setFontWeight('bold').setBackground('#d9ead3');
  cat.getRange(1, revCol1).setValue(SQ_.REVENUE_COL_NAME);
  cat.getRange(1, revCol1).setFontWeight('bold').setBackground('#cfe2f3');

  // 從 Tally 彙總數量
  var sold = {};
  mergeSold_(sold, sumFromTickets_(ss));
  mergeSold_(sold, sumFromMerch_(ss));

  var data = rangeInc_(cat, 2, 1, lastRow, lastCol).getValues();
  var outK = [];
  var outL = [];
  var updated = 0;
  var revenueSum = 0;
  var r;
  for (r = 0; r < data.length; r++) {
    var level = Number(data[r][levelCol0]);
    if (level !== 3) {
      outK.push(['']);
      outL.push(['']);
      continue;
    }

    var keys = buildMatchKeys_(data[r], idx);
    var qty = 0;
    var k;
    for (k = 0; k < keys.length; k++) {
      var key = keys[k];
      if (key && sold[key] != null) qty += sold[key];
    }
    if (idx.sku != null) {
      var sku = normKey_(data[r][idx.sku]);
      if (sku && sold[sku] != null) qty += sold[sku];
    }

    var price = toMoney_(data[r][priceCol0]);
    var revenue = qty * price;

    outK.push([qty]);
    outL.push([revenue]);
    revenueSum += revenue;
    updated++;
  }

  // 只寫 K、L 欄
  rangeInc_(cat, 2, qtyCol1, lastRow, qtyCol1).setValues(outK);
  rangeInc_(cat, 2, revCol1, lastRow, revCol1).setValues(outL);

  var lines = [];
  var keys2 = Object.keys(sold);
  keys2.sort();
  var i;
  for (i = 0; i < Math.min(keys2.length, 15); i++) {
    lines.push(keys2[i] + ' → ' + sold[keys2[i]]);
  }

  SpreadsheetApp.getUi().alert(
    'Order_Categories 已更新\n' +
    'K 欄 = 銷售數量（Tally 加總）\n' +
    'L 欄 = 單件貨品總收入（數量 × 售價）\n' +
    'SKU 列: ' + updated + '\n' +
    '總收入合計: HK$' + revenueSum + '\n\n' +
    (lines.length ? lines.join('\n') : '（尚無 Tally 銷售資料）')
  );
}

function toMoney_(v) {
  if (v === '' || v == null) return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var n = Number(String(v).replace(/[^0-9.\-]/g, ''));
  return isFinite(n) ? n : 0;
}

/* ========== 彙總 NA_Tickets ========== */

function sumFromTickets_(ss) {
  var sold = {};
  var sh = ss.getSheetByName(SQ_.TICKETS);
  if (!sh || sh.getLastRow() < 2) return sold;

  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  var headers = rangeInc_(sh, 1, 1, 1, lastCol).getValues()[0];
  var data = rangeInc_(sh, 2, 1, lastRow, lastCol).getValues();
  var hmap = headerIndexMap_(headers);

  // 新格式：門票類別 + 數量
  var typeCol = firstCol_(hmap, ['門票類別', '票種', 'Ticket Type']);
  var qtyCol = firstCol_(hmap, ['數量', 'Qty', 'qty']);

  var r, c;
  if (typeCol != null && qtyCol != null) {
    for (r = 0; r < data.length; r++) {
      var sid = String(data[r][0] || '');
      // 可選：不計虛擬 TEST- 單？預設要計入測試銷量，方便驗證
      var type = String(data[r][typeCol] || '').trim();
      var qty = toQty_(data[r][qtyCol]);
      if (!type || qty <= 0) continue;
      addSold_(sold, type, qty);
      addSold_(sold, stripEmoji_(type), qty);
    }
    return sold;
  }

  // 舊格式：每一票種一欄數量
  for (c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim();
    if (!h) continue;
    if (!isTicketHeader_(h)) continue;
    for (r = 0; r < data.length; r++) {
      var q = toQty_(data[r][c]);
      if (q > 0) {
        addSold_(sold, h, q);
        addSold_(sold, stripEmoji_(h), q);
      }
    }
  }
  return sold;
}

/* ========== 彙總 NA_Merch ========== */

function sumFromMerch_(ss) {
  var sold = {};
  var sh = ss.getSheetByName(SQ_.MERCH);
  if (!sh || sh.getLastRow() < 2) return sold;

  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  var headers = rangeInc_(sh, 1, 1, 1, lastCol).getValues()[0];
  var data = rangeInc_(sh, 2, 1, lastRow, lastCol).getValues();
  var hmap = headerIndexMap_(headers);

  var catCol = firstCol_(hmap, ['商品分欄', '商品類別', '商品', 'Merch Category']);
  var qtyCol = firstCol_(hmap, ['數量', 'Qty', 'qty']);

  // 新格式：商品分欄 + 數量（分欄可能是「服飾」等細類，或品名）
  var r, c;
  if (catCol != null && qtyCol != null) {
    for (r = 0; r < data.length; r++) {
      var cat = String(data[r][catCol] || '').trim();
      var qty = toQty_(data[r][qtyCol]);
      if (!cat || qty <= 0) continue;
      addSold_(sold, cat, qty);
      // 細類 → 可能對多個 SKU：另外用細類 key 累加，SKU 端用 sub_category_zh 對
      addSold_(sold, 'sub:' + cat, qty);
    }
  }

  // 舊格式／並存：各商品數量欄（Ark T-Shirt…）
  for (c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim();
    if (!h) continue;
    if (c === catCol || c === qtyCol) continue;
    if (!isMerchProductHeader_(h)) continue;
    for (r = 0; r < data.length; r++) {
      var q = toQty_(data[r][c]);
      if (q > 0) {
        addSold_(sold, h, q);
        addSold_(sold, stripEmoji_(h), q);
      }
    }
  }
  return sold;
}

/* ========== 寫入對應 ========== */

function buildMatchKeys_(row, idx) {
  var keys = [];
  var fields = [
    'form_option_label',
    'tally_column_hint',
    'name_zh',
    'name_en',
    'sku',
    'sub_category_zh',
    'sub_category'
  ];
  var i;
  for (i = 0; i < fields.length; i++) {
    var f = fields[i];
    if (idx[f] == null) continue;
    var v = String(row[idx[f]] || '').trim();
    if (!v) continue;
    keys.push(normKey_(v));
    keys.push(normKey_(stripEmoji_(v)));
    if (f === 'sub_category_zh' || f === 'sub_category') {
      keys.push(normKey_('sub:' + v));
    }
  }
  return keys;
}

function addSold_(sold, key, qty) {
  var k = normKey_(key);
  if (!k) return;
  sold[k] = (sold[k] || 0) + qty;
}

function mergeSold_(a, b) {
  var keys = Object.keys(b);
  var i;
  for (i = 0; i < keys.length; i++) {
    var k = keys[i];
    a[k] = (a[k] || 0) + b[k];
  }
}

function headerIndexMap_(headers) {
  var m = {};
  var i;
  for (i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').trim();
    if (h) m[h] = i;
  }
  return m;
}

function firstCol_(hmap, names) {
  var i;
  for (i = 0; i < names.length; i++) {
    if (hmap[names[i]] != null) return hmap[names[i]];
  }
  // 模糊
  var keys = Object.keys(hmap);
  for (i = 0; i < names.length; i++) {
    var n = names[i];
    var k;
    for (k = 0; k < keys.length; k++) {
      if (keys[k].indexOf(n) >= 0) return hmap[keys[k]];
    }
  }
  return null;
}

function isTicketHeader_(h) {
  return /早鳥|預售|會員特級|Metal Pass|Early|Advanced|門票|HKD\s*3/i.test(h);
}

function isMerchProductHeader_(h) {
  if (/Submission|Respondent|Submitted|Total|Name|Email|Tel|Metal Pass|付款|Payment|數量|單價|商品分欄|ID/i.test(h)) {
    return false;
  }
  return /Ark |T-Shirt|Tower|Keychain|Pack|Tote|Patch|Mousepad|Tee|毛巾|鎖匙|布章/i.test(h);
}

function toQty_(v) {
  if (v === '' || v == null) return 0;
  var n = Number(v);
  if (!isFinite(n)) return 0;
  return Math.max(0, Math.floor(n));
}

function stripEmoji_(s) {
  return String(s || '').replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, '').trim();
}

function normKey_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, '')
    .replace(/\s+/g, '')
    .replace(/[^\w\u4e00-\u9fff]/g, '');
}
