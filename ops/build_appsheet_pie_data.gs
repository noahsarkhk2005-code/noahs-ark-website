/**
 * 從 NA_Tickets / NA_Merch 整理 AppSheet 圓餅圖資料庫
 *
 * 產出分頁：
 *   AS_Pie_Tickets  — 票務：門票類型 / 總銷量 / 總銷售額
 *   AS_Pie_Merch    — 商品：貨品／分欄 / 總銷量 / 總銷售額
 *   AS_Pie_All      — 合併（可選：全部 SKU 一表）
 *
 * AppSheet 用法：
 *   Pie（票務·銷售額）:  label=ticket_type, value=revenue
 *   Pie（票務·銷量）:    label=ticket_type, value=sales_qty
 *   Pie（商品·銷售額）:  label=product_label, value=revenue
 *   Pie（商品·銷量）:    label=product_label, value=sales_qty
 *
 * 執行：buildAppSheetPieData
 */

var PIE_ = {
  TICKETS_SRC: 'NA_Tickets',
  MERCH_SRC: 'NA_Merch',
  CAT: 'Order_Categories',
  OUT_TICKETS: 'AS_Pie_Tickets',
  OUT_MERCH: 'AS_Pie_Merch',
  OUT_ALL: 'AS_Pie_All',
  // 後備單價（Order_Categories 對不到時）
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

function rangeInc_(sh, r1, c1, r2, c2) {
  return sh.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1);
}

function buildAppSheetPieData() {
  var ss = SpreadsheetApp.getActive();
  var priceMap = loadPriceMapFromCategories_(ss);

  var ticketAgg = aggregateTickets_(ss, priceMap);
  var merchAgg = aggregateMerch_(ss, priceMap);

  writePieTickets_(ss, ticketAgg);
  writePieMerch_(ss, merchAgg);
  writePieAll_(ss, ticketAgg, merchAgg);

  SpreadsheetApp.getUi().alert(
    'AppSheet 圓餅圖資料已更新\n\n' +
    'AS_Pie_Tickets（票務）: ' + ticketAgg.length + ' 種門票\n' +
    'AS_Pie_Merch（商品）: ' + merchAgg.length + ' 項\n' +
    'AS_Pie_All（合併）: ' + (ticketAgg.length + merchAgg.length) + ' 列\n\n' +
    'AppSheet → Data → 加入這三表 → Sync\n' +
    'Pie Chart：Category = *_label / ticket_type / product_label\n' +
    '          Value = revenue 或 sales_qty'
  );
}

/* ========== 彙總 ========== */

function aggregateTickets_(ss, priceMap) {
  var sh = ss.getSheetByName(PIE_.TICKETS_SRC);
  var map = {}; // label -> {qty, revenue}
  if (!sh || sh.getLastRow() < 2) return [];

  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  var headers = rangeInc_(sh, 1, 1, 1, lastCol).getValues()[0];
  var data = rangeInc_(sh, 2, 1, lastRow, lastCol).getValues();
  var hmap = headerIndexMap_(headers);

  var typeCol = firstCol_(hmap, ['門票類別', '票種', 'Ticket Type']);
  var qtyCol = firstCol_(hmap, ['數量', 'Qty', 'qty']);
  var priceCol = firstCol_(hmap, ['單價', 'Price', 'list_price']);

  var r, c;
  if (typeCol != null && qtyCol != null) {
    for (r = 0; r < data.length; r++) {
      var type = String(data[r][typeCol] || '').trim();
      var qty = toQty_(data[r][qtyCol]);
      if (!type || qty <= 0) continue;
      var unit = priceCol != null ? toMoney_(data[r][priceCol]) : 0;
      if (!unit) unit = lookupPrice_(type, priceMap, PIE_.TICKET_PRICES);
      addAgg_(map, type, qty, unit * qty);
    }
  } else {
    // 舊：每票種一欄
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
  return mapToRows_(map, 'ticket');
}

function aggregateMerch_(ss, priceMap) {
  var sh = ss.getSheetByName(PIE_.MERCH_SRC);
  var map = {};
  if (!sh || sh.getLastRow() < 2) return [];

  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  var headers = rangeInc_(sh, 1, 1, 1, lastCol).getValues()[0];
  var data = rangeInc_(sh, 2, 1, lastRow, lastCol).getValues();
  var hmap = headerIndexMap_(headers);

  var catCol = firstCol_(hmap, ['商品分欄', '商品類別', 'Merch Category']);
  var qtyCol = firstCol_(hmap, ['數量', 'Qty', 'qty']);
  var priceCol = firstCol_(hmap, ['單價', 'Price', 'list_price']);

  var r, c;
  // 新格式 IJK
  if (catCol != null && qtyCol != null) {
    for (r = 0; r < data.length; r++) {
      var cat = String(data[r][catCol] || '').trim();
      var qty = toQty_(data[r][qtyCol]);
      if (!cat || qty <= 0) continue;
      var unit = priceCol != null ? toMoney_(data[r][priceCol]) : 0;
      if (!unit) unit = lookupPrice_(cat, priceMap, PIE_.MERCH_PRICES);
      addAgg_(map, cat, qty, unit * qty);
    }
  }

  // 舊／並存：各商品數量欄
  for (c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim();
    if (!h) continue;
    if (c === catCol || c === qtyCol || c === priceCol) continue;
    if (!isMerchProductHeader_(h)) continue;
    for (r = 0; r < data.length; r++) {
      var q = toQty_(data[r][c]);
      if (q <= 0) continue;
      var u = lookupPrice_(h, priceMap, PIE_.MERCH_PRICES);
      if (!u) u = guessPriceFromHeader_(h);
      addAgg_(map, h, q, u * q);
    }
  }

  return mapToRows_(map, 'merch');
}

/* ========== 寫出分頁 ========== */

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
      'TKT-' + (i + 1),
      'ticket',
      x.label,
      x.label,
      x.qty,
      x.revenue,
      avg,
      new Date()
    ]);
  }
  writeSheet_(ss, PIE_.OUT_TICKETS, headers, body, '#fce8e6');
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
      'MRC-' + (i + 1),
      'merch',
      x.label,
      x.label,
      x.qty,
      x.revenue,
      avg,
      new Date()
    ]);
  }
  writeSheet_(ss, PIE_.OUT_MERCH, headers, body, '#e6f4ea');
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
  writeSheet_(ss, PIE_.OUT_ALL, headers, body, '#fff2cc');
}

function writeSheet_(ss, name, headers, body, headerColor) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground(headerColor);
  sh.setFrozenRows(1);
  if (body.length) {
    sh.getRange(2, 1, body.length, headers.length).setValues(body);
  }
  // 數字格式
  var revCol = headers.indexOf('revenue') + 1;
  var qtyCol = headers.indexOf('sales_qty') + 1;
  if (body.length && revCol > 0) {
    sh.getRange(2, revCol, body.length, 1).setNumberFormat('"$"#,##0');
  }
  if (body.length && qtyCol > 0) {
    sh.getRange(2, qtyCol, body.length, 1).setNumberFormat('#,##0');
  }
}

/* ========== helpers ========== */

function loadPriceMapFromCategories_(ss) {
  var map = {};
  var sh = ss.getSheetByName(PIE_.CAT);
  if (!sh || sh.getLastRow() < 2) return map;
  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  var headers = rangeInc_(sh, 1, 1, 1, lastCol).getValues()[0];
  var idx = headerIndexMap_(headers);
  var data = rangeInc_(sh, 2, 1, lastRow, lastCol).getValues();
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
  // fallback exact
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

function mapToRows_(map, channel) {
  var keys = Object.keys(map);
  keys.sort();
  var out = [];
  var i;
  for (i = 0; i < keys.length; i++) {
    out.push(map[keys[i]]);
  }
  // 銷售額高 → 前
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
