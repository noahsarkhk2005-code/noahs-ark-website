/**
 * 從 NA_Tickets / NA_Merch 整理 AppSheet 圓餅圖資料
 * 已優化：避免 clear 整表逾時、分表寫入、單次 setValues
 *
 * 執行：
 *   buildAppSheetPieData     — 三表一次（建議）
 *   buildPieTicketsOnly      — 只寫票務（逾時可分段跑）
 *   buildPieMerchOnly        — 只寫商品
 */

var PIE_ = {
  TICKETS_SRC: 'NA_Tickets',
  MERCH_SRC: 'NA_Merch',
  CAT: 'Order_Categories',
  OUT_TICKETS: 'AS_Pie_Tickets',
  OUT_MERCH: 'AS_Pie_Merch',
  OUT_ALL: 'AS_Pie_All',
  TICKET_PRICES: {
    '會員特級優惠 HKD300': 300,
    '🎫 早鳥門票 HKD300': 300,
    '早鳥門票 HKD300': 300,
    '🎫 預售 HKD350': 350,
    '預售 HKD350': 350
  },
  MERCH_PRICES: {
    '服飾': 280,
    '配件': 120,
    '毛巾': 80,
    '套裝 Pack': 200,
    '袋類': 120,
    'Ark T-Shirt HKD280': 280,
    'Ark Tower/毛巾 HKD80': 80,
    'Ark Keychain/鎖匙扣 HKD120': 120,
    'Ark BigPack HKD200': 200,
    'Ark TinyPack HKD60': 60,
    'Ark Tote Bag HKD120': 120
  }
};

/** row, col, numRows, numCols — Apps Script 正確用法 */
function rg_(sh, r, c, nR, nC) {
  return sh.getRange(r, c, nR, nC);
}

function buildAppSheetPieData() {
  var ss = SpreadsheetApp.getActive();
  var priceMap = loadPriceMapFromCategories_(ss);

  var ticketAgg = aggregateTickets_(ss, priceMap);
  SpreadsheetApp.flush();
  writePieTickets_(ss, ticketAgg);
  SpreadsheetApp.flush();

  var merchAgg = aggregateMerch_(ss, priceMap);
  SpreadsheetApp.flush();
  writePieMerch_(ss, merchAgg);
  SpreadsheetApp.flush();

  writePieAll_(ss, ticketAgg, merchAgg);
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    'AppSheet 圓餅圖資料已更新\n\n' +
    'AS_Pie_Tickets: ' + ticketAgg.length + ' 種\n' +
    'AS_Pie_Merch: ' + merchAgg.length + ' 項\n' +
    'AS_Pie_All: ' + (ticketAgg.length + merchAgg.length) + ' 列\n\n' +
    '若再逾時請改跑 buildPieTicketsOnly / buildPieMerchOnly'
  );
}

function buildPieTicketsOnly() {
  var ss = SpreadsheetApp.getActive();
  var priceMap = loadPriceMapFromCategories_(ss);
  var rows = aggregateTickets_(ss, priceMap);
  writePieTickets_(ss, rows);
  SpreadsheetApp.getUi().alert('AS_Pie_Tickets 完成：' + rows.length + ' 種門票');
}

function buildPieMerchOnly() {
  var ss = SpreadsheetApp.getActive();
  var priceMap = loadPriceMapFromCategories_(ss);
  var rows = aggregateMerch_(ss, priceMap);
  writePieMerch_(ss, rows);
  SpreadsheetApp.getUi().alert('AS_Pie_Merch 完成：' + rows.length + ' 項商品');
}

/* ========== 彙總（只讀需要的列欄，跳過空列） ========== */

function aggregateTickets_(ss, priceMap) {
  var sh = ss.getSheetByName(PIE_.TICKETS_SRC);
  var map = {};
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  // 限制最大欄，避免讀到過寬範圍
  lastCol = Math.min(lastCol, 20);
  var headers = rg_(sh, 1, 1, 1, lastCol).getValues()[0];
  var numData = lastRow - 1;
  var data = rg_(sh, 2, 1, numData, lastCol).getValues();
  var hmap = headerIndexMap_(headers);

  var typeCol = firstCol_(hmap, ['門票類別', '票種', 'Ticket Type']);
  var qtyCol = firstCol_(hmap, ['數量', 'Qty', 'qty']);
  var priceCol = firstCol_(hmap, ['單價', 'Price', 'list_price']);

  var r, c;
  if (typeCol != null && qtyCol != null) {
    for (r = 0; r < data.length; r++) {
      var type = String(data[r][typeCol] || '').trim();
      if (!type) continue;
      var qty = toQty_(data[r][qtyCol]);
      if (qty <= 0) continue;
      var unit = priceCol != null ? toMoney_(data[r][priceCol]) : 0;
      if (!unit) unit = lookupPrice_(type, priceMap, PIE_.TICKET_PRICES);
      addAgg_(map, type, qty, unit * qty);
    }
  } else {
    for (c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').trim();
      if (!isTicketHeader_(h)) continue;
      for (r = 0; r < data.length; r++) {
        var q = toQty_(data[r][c]);
        if (q <= 0) continue;
        var u = lookupPrice_(h, priceMap, PIE_.TICKET_PRICES);
        addAgg_(map, h, q, u * q);
      }
    }
  }
  return mapToRows_(map);
}

function aggregateMerch_(ss, priceMap) {
  var sh = ss.getSheetByName(PIE_.MERCH_SRC);
  var map = {};
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  lastCol = Math.min(lastCol, 25);
  var headers = rg_(sh, 1, 1, 1, lastCol).getValues()[0];
  var numData = lastRow - 1;
  var data = rg_(sh, 2, 1, numData, lastCol).getValues();
  var hmap = headerIndexMap_(headers);

  var catCol = firstCol_(hmap, ['商品分欄', '商品類別', 'Merch Category']);
  var qtyCol = firstCol_(hmap, ['數量', 'Qty', 'qty']);
  var priceCol = firstCol_(hmap, ['單價', 'Price', 'list_price']);

  var r, c;
  if (catCol != null && qtyCol != null) {
    for (r = 0; r < data.length; r++) {
      var cat = String(data[r][catCol] || '').trim();
      if (!cat) continue;
      var qty = toQty_(data[r][qtyCol]);
      if (qty <= 0) continue;
      var unit = priceCol != null ? toMoney_(data[r][priceCol]) : 0;
      if (!unit) unit = lookupPrice_(cat, priceMap, PIE_.MERCH_PRICES);
      addAgg_(map, cat, qty, unit * qty);
    }
  }

  // 舊商品欄：只掃一次 headers 中的商品欄
  var productCols = [];
  for (c = 0; c < headers.length; c++) {
    if (c === catCol || c === qtyCol || c === priceCol) continue;
    var hh = String(headers[c] || '').trim();
    if (isMerchProductHeader_(hh)) productCols.push({ c: c, h: hh });
  }
  for (r = 0; r < data.length; r++) {
    for (c = 0; c < productCols.length; c++) {
      var pc = productCols[c];
      var q = toQty_(data[r][pc.c]);
      if (q <= 0) continue;
      var u = lookupPrice_(pc.h, priceMap, PIE_.MERCH_PRICES);
      if (!u) u = guessPriceFromHeader_(pc.h);
      addAgg_(map, pc.h, q, u * q);
    }
  }

  return mapToRows_(map);
}

/* ========== 寫出（避免 timeout） ========== */

function writePieTickets_(ss, rows) {
  var headers = [
    'row_id', 'channel', 'ticket_type', 'chart_label',
    'sales_qty', 'revenue', 'unit_price_avg', 'updated_at'
  ];
  var body = [];
  var i;
  for (i = 0; i < rows.length; i++) {
    var x = rows[i];
    var avg = x.qty > 0 ? Math.round((x.revenue / x.qty) * 100) / 100 : 0;
    body.push([
      'TKT-' + (i + 1), 'ticket', x.label, x.label,
      x.qty, x.revenue, avg, new Date()
    ]);
  }
  writeSheetFast_(ss, PIE_.OUT_TICKETS, headers, body, '#fce8e6');
}

function writePieMerch_(ss, rows) {
  var headers = [
    'row_id', 'channel', 'product_label', 'chart_label',
    'sales_qty', 'revenue', 'unit_price_avg', 'updated_at'
  ];
  var body = [];
  var i;
  for (i = 0; i < rows.length; i++) {
    var x = rows[i];
    var avg = x.qty > 0 ? Math.round((x.revenue / x.qty) * 100) / 100 : 0;
    body.push([
      'MRC-' + (i + 1), 'merch', x.label, x.label,
      x.qty, x.revenue, avg, new Date()
    ]);
  }
  writeSheetFast_(ss, PIE_.OUT_MERCH, headers, body, '#e6f4ea');
}

function writePieAll_(ss, tickets, merch) {
  var headers = [
    'row_id', 'channel', 'channel_zh', 'item_label', 'chart_label',
    'sales_qty', 'revenue', 'unit_price_avg', 'updated_at'
  ];
  var body = [];
  var i;
  for (i = 0; i < tickets.length; i++) {
    var t = tickets[i];
    var avgT = t.qty > 0 ? Math.round((t.revenue / t.qty) * 100) / 100 : 0;
    body.push([
      'ALL-T-' + (i + 1), 'ticket', '票務', t.label, t.label,
      t.qty, t.revenue, avgT, new Date()
    ]);
  }
  for (i = 0; i < merch.length; i++) {
    var m = merch[i];
    var avgM = m.qty > 0 ? Math.round((m.revenue / m.qty) * 100) / 100 : 0;
    body.push([
      'ALL-M-' + (i + 1), 'merch', '商品', m.label, m.label,
      m.qty, m.revenue, avgM, new Date()
    ]);
  }
  writeSheetFast_(ss, PIE_.OUT_ALL, headers, body, '#fff2cc');
}

/**
 * 快速寫入：不清空整張超大表
 * - 只 clearContent 實際用到的範圍
 * - 一次 setValues 寫入 header+body
 * - 不刪除分頁（保留 AppSheet sheet id）
 */
function writeSheetFast_(ss, name, headers, body, headerColor) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }

  var cols = headers.length;
  var rows = 1 + body.length; // header + data
  // 至少寫 1 列表頭
  if (rows < 1) rows = 1;

  // 確保有足夠列/欄（必要時擴充，但不要亂刪）
  var needExtraRows = rows - sh.getMaxRows();
  if (needExtraRows > 0) sh.insertRowsAfter(sh.getMaxRows(), needExtraRows);
  var needExtraCols = cols - sh.getMaxColumns();
  if (needExtraCols > 0) sh.insertColumnsAfter(sh.getMaxColumns(), needExtraCols);

  // 只清舊資料區（最多清到「舊 lastRow」與「新 rows」的較大者，且封頂 200 列）
  var oldLast = Math.min(Math.max(sh.getLastRow(), rows), 200);
  var oldCols = Math.min(Math.max(sh.getLastColumn(), cols), 20);
  if (oldLast >= 1 && oldCols >= 1) {
    try {
      rg_(sh, 1, 1, oldLast, oldCols).clearContent();
    } catch (e) {
      // 若 clear 仍 timeout，改直接覆寫
    }
  }

  // 組完整二維陣列：第一列 header
  var all = [headers];
  var i;
  for (i = 0; i < body.length; i++) {
    all.push(body[i]);
  }

  // 一次寫入（小資料量：pie 通常 < 50 列）
  rg_(sh, 1, 1, all.length, cols).setValues(all);

  // 輕量格式（只表頭）
  try {
    rg_(sh, 1, 1, 1, cols).setFontWeight('bold').setBackground(headerColor);
    sh.setFrozenRows(1);
  } catch (e2) {
    // ignore format errors
  }
}

/* ========== helpers ========== */

function loadPriceMapFromCategories_(ss) {
  var map = {};
  var sh = ss.getSheetByName(PIE_.CAT);
  if (!sh) return map;
  var lastRow = sh.getLastRow();
  var lastCol = Math.min(sh.getLastColumn(), 20);
  if (lastRow < 2 || lastCol < 1) return map;

  var headers = rg_(sh, 1, 1, 1, lastCol).getValues()[0];
  var idx = headerIndexMap_(headers);
  if (idx.level == null || idx.list_price == null) return map;

  var data = rg_(sh, 2, 1, lastRow - 1, lastCol).getValues();
  var r;
  for (r = 0; r < data.length; r++) {
    if (Number(data[r][idx.level]) !== 3) continue;
    var price = toMoney_(data[r][idx.list_price]);
    if (!price) continue;
    var fields = ['form_option_label', 'tally_column_hint', 'name_zh', 'name_en', 'sku', 'sub_category_zh'];
    var f;
    for (f = 0; f < fields.length; f++) {
      var key = fields[f];
      if (idx[key] == null) continue;
      var v = String(data[r][idx[key]] || '').trim();
      if (v) map[normKey_(v)] = price;
    }
  }
  return map;
}

function lookupPrice_(label, priceMap, fallback) {
  var k = normKey_(label);
  if (priceMap[k] != null) return priceMap[k];
  if (fallback[label] != null) return fallback[label];
  var keys = Object.keys(fallback);
  var i;
  for (i = 0; i < keys.length; i++) {
    if (normKey_(keys[i]) === k) return fallback[keys[i]];
  }
  return guessPriceFromHeader_(label);
}

function addAgg_(map, label, qty, revenue) {
  var L = String(label || '').trim();
  if (!L) return;
  if (!map[L]) map[L] = { label: L, qty: 0, revenue: 0 };
  map[L].qty += qty;
  map[L].revenue += revenue;
}

function mapToRows_(map) {
  var keys = Object.keys(map);
  var out = [];
  var i;
  for (i = 0; i < keys.length; i++) out.push(map[keys[i]]);
  out.sort(function (a, b) { return b.revenue - a.revenue; });
  return out;
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
  var i, k;
  for (i = 0; i < names.length; i++) {
    if (hmap[names[i]] != null) return hmap[names[i]];
  }
  var keys = Object.keys(hmap);
  for (i = 0; i < names.length; i++) {
    for (k = 0; k < keys.length; k++) {
      if (keys[k].indexOf(names[i]) >= 0) return hmap[keys[k]];
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

function guessPriceFromHeader_(h) {
  var m = String(h || '').match(/HKD?\s*(\d+)/i);
  return m ? Number(m[1]) : 0;
}

function toQty_(v) {
  if (v === '' || v == null) return 0;
  var n = Number(v);
  return isFinite(n) ? Math.max(0, Math.floor(n)) : 0;
}

function toMoney_(v) {
  if (v === '' || v == null) return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var n = Number(String(v).replace(/[^0-9.\-]/g, ''));
  return isFinite(n) ? n : 0;
}

function normKey_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, '')
    .replace(/\s+/g, '')
    .replace(/[^\w\u4e00-\u9fff]/g, '');
}
