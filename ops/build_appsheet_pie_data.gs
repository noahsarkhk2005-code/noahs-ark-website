/**
 * AppSheet 圓餅圖資料
 *
 * 分類顯示名稱 = Order_Categories.tally_column_hint
 * （沒有 hint 則用 form_option_label → name_zh）
 *
 * 數值：
 *   sales_qty  = Tally 訂單該品項數量加總
 *   revenue    = 數量 × list_price
 *
 * 產出：
 *   AS_Pie_Tickets — channel=ticket，label=tally_column_hint
 *   AS_Pie_Merch   — channel=merch
 *   AS_Pie_All     — 合併
 *
 * 執行：buildAppSheetPieData
 * 分段：buildPieTicketsOnly / buildPieMerchOnly
 */

var PIE_ = {
  TICKETS_SRC: 'NA_Tickets',
  MERCH_SRC: 'NA_Merch',
  CAT: 'Order_Categories',
  OUT_TICKETS: 'AS_Pie_Tickets',
  OUT_MERCH: 'AS_Pie_Merch',
  OUT_ALL: 'AS_Pie_All'
};

function rg_(sh, r, c, nR, nC) {
  return sh.getRange(r, c, nR, nC);
}

function buildAppSheetPieData() {
  var ss = SpreadsheetApp.getActive();
  var catalog = loadCatalogSkus_(ss); // [{channel, displayLabel, price, matchKeys[]}]
  var sold = loadSoldFromTally_(ss);  // normKey -> qty

  var tickets = buildPieRows_(catalog, sold, 'ticket');
  var merch = buildPieRows_(catalog, sold, 'merch');

  SpreadsheetApp.flush();
  writePieTickets_(ss, tickets);
  SpreadsheetApp.flush();
  writePieMerch_(ss, merch);
  SpreadsheetApp.flush();
  writePieAll_(ss, tickets, merch);
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    'Pie 資料已更新（分類 = tally_column_hint）\n\n' +
    '票務: ' + tickets.length + ' 項\n' +
    '商品: ' + merch.length + ' 項\n\n' +
    'AppSheet Pie:\n' +
    '  Label = chart_label（= tally_column_hint）\n' +
    '  Value = revenue 或 sales_qty'
  );
}

function buildPieTicketsOnly() {
  var ss = SpreadsheetApp.getActive();
  var catalog = loadCatalogSkus_(ss);
  var sold = loadSoldFromTally_(ss);
  var tickets = buildPieRows_(catalog, sold, 'ticket');
  writePieTickets_(ss, tickets);
  SpreadsheetApp.getUi().alert('AS_Pie_Tickets: ' + tickets.length + '（label=tally_column_hint）');
}

function buildPieMerchOnly() {
  var ss = SpreadsheetApp.getActive();
  var catalog = loadCatalogSkus_(ss);
  var sold = loadSoldFromTally_(ss);
  var merch = buildPieRows_(catalog, sold, 'merch');
  writePieMerch_(ss, merch);
  SpreadsheetApp.getUi().alert('AS_Pie_Merch: ' + merch.length + '（label=tally_column_hint）');
}

/* ========== 主檔：Order_Categories ========== */

/**
 * 每個 level=3 SKU 一筆
 * displayLabel 優先：tally_column_hint → form_option_label → name_zh → sku
 */
function loadCatalogSkus_(ss) {
  var sh = ss.getSheetByName(PIE_.CAT);
  if (!sh) throw new Error('找不到 Order_Categories');
  var lastRow = sh.getLastRow();
  var lastCol = Math.min(sh.getLastColumn(), 25);
  if (lastRow < 2) return [];

  var headers = rg_(sh, 1, 1, 1, lastCol).getValues()[0];
  var idx = headerIndexMap_(headers);
  var data = rg_(sh, 2, 1, lastRow - 1, lastCol).getValues();
  var out = [];
  var r;
  for (r = 0; r < data.length; r++) {
    if (Number(data[r][idx.level]) !== 3) continue;
    if (idx.is_active != null) {
      var act = String(data[r][idx.is_active]).toUpperCase();
      if (act === 'FALSE' || act === '0' || act === 'NO') continue;
    }

    var channel = String(data[r][idx.channel] || '').trim().toLowerCase();
    if (channel !== 'ticket' && channel !== 'merch') continue;

    var hint = cell_(data[r], idx, 'tally_column_hint');
    var form = cell_(data[r], idx, 'form_option_label');
    var nameZh = cell_(data[r], idx, 'name_zh');
    var nameEn = cell_(data[r], idx, 'name_en');
    var sku = cell_(data[r], idx, 'sku');
    var subZh = cell_(data[r], idx, 'sub_category_zh');
    var price = toMoney_(idx.list_price != null ? data[r][idx.list_price] : 0);

    // ★ 分類顯示名稱 = tally_column_hint
    var displayLabel = hint || form || nameZh || sku;
    if (!displayLabel) continue;
    if (!price) price = guessPriceFromHeader_(displayLabel);

    // 比對 Tally 用的 keys（多別名）
    var matchKeys = [];
    pushKey_(matchKeys, hint);
    pushKey_(matchKeys, form);
    pushKey_(matchKeys, nameZh);
    pushKey_(matchKeys, nameEn);
    pushKey_(matchKeys, sku);
    pushKey_(matchKeys, subZh);
    pushKey_(matchKeys, stripEmoji_(hint));
    pushKey_(matchKeys, stripEmoji_(form));

    out.push({
      channel: channel,
      displayLabel: displayLabel, // = tally_column_hint 優先
      tally_column_hint: hint,
      sku: sku,
      price: price,
      matchKeys: matchKeys
    });
  }
  return out;
}

function cell_(row, idx, name) {
  if (idx[name] == null) return '';
  return String(row[idx[name]] || '').trim();
}

function pushKey_(arr, v) {
  var k = normKey_(v);
  if (!k) return;
  if (arr.indexOf(k) < 0) arr.push(k);
}

/* ========== Tally 原始銷量 ========== */

function loadSoldFromTally_(ss) {
  var sold = {};
  mergeSold_(sold, sumFromTicketsRaw_(ss));
  mergeSold_(sold, sumFromMerchRaw_(ss));
  return sold;
}

function sumFromTicketsRaw_(ss) {
  var sold = {};
  var sh = ss.getSheetByName(PIE_.TICKETS_SRC);
  if (!sh || sh.getLastRow() < 2) return sold;
  var lastCol = Math.min(sh.getLastColumn(), 20);
  var lastRow = sh.getLastRow();
  var headers = rg_(sh, 1, 1, 1, lastCol).getValues()[0];
  var data = rg_(sh, 2, 1, lastRow - 1, lastCol).getValues();
  var hmap = headerIndexMap_(headers);

  var typeCol = firstCol_(hmap, ['門票類別', '票種', 'Ticket Type']);
  var qtyCol = firstCol_(hmap, ['數量', 'Qty', 'qty']);

  var r, c;
  if (typeCol != null && qtyCol != null) {
    for (r = 0; r < data.length; r++) {
      var type = String(data[r][typeCol] || '').trim();
      var qty = toQty_(data[r][qtyCol]);
      if (!type || qty <= 0) continue;
      addSold_(sold, type, qty);
      addSold_(sold, stripEmoji_(type), qty);
    }
  }
  // 舊欄：每票種一欄（欄名常 = tally_column_hint）
  for (c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim();
    if (!h || !isTicketHeader_(h)) continue;
    if (c === typeCol || c === qtyCol) continue;
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

function sumFromMerchRaw_(ss) {
  var sold = {};
  var sh = ss.getSheetByName(PIE_.MERCH_SRC);
  if (!sh || sh.getLastRow() < 2) return sold;
  var lastCol = Math.min(sh.getLastColumn(), 25);
  var lastRow = sh.getLastRow();
  var headers = rg_(sh, 1, 1, 1, lastCol).getValues()[0];
  var data = rg_(sh, 2, 1, lastRow - 1, lastCol).getValues();
  var hmap = headerIndexMap_(headers);

  var catCol = firstCol_(hmap, ['商品分欄', '商品類別']);
  var qtyCol = firstCol_(hmap, ['數量', 'Qty', 'qty']);

  var r, c;
  if (catCol != null && qtyCol != null) {
    for (r = 0; r < data.length; r++) {
      var cat = String(data[r][catCol] || '').trim();
      var qty = toQty_(data[r][qtyCol]);
      if (!cat || qty <= 0) continue;
      addSold_(sold, cat, qty);
      addSold_(sold, stripEmoji_(cat), qty);
    }
  }
  // 舊欄：欄名 = tally_column_hint（Ark T-Shirt HKD280…）
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

/* ========== 組 pie 列：每個 tally_column_hint 一列 ========== */

function buildPieRows_(catalog, sold, channel) {
  // 同一 tally_column_hint（displayLabel）合併
  // 別名 key（emoji/無 emoji）只取 max，避免同一筆銷量被加兩次
  var byLabel = {};
  var i, j, k;

  for (i = 0; i < catalog.length; i++) {
    var item = catalog[i];
    if (item.channel !== channel) continue;

    var label = item.displayLabel;
    if (!byLabel[label]) {
      byLabel[label] = {
        label: label,
        tally_column_hint: item.tally_column_hint || label,
        sku: item.sku,
        price: item.price || 0,
        keys: {}
      };
    }
    if (item.price) byLabel[label].price = item.price;
    for (j = 0; j < item.matchKeys.length; j++) {
      byLabel[label].keys[item.matchKeys[j]] = true;
    }
  }

  var labels = Object.keys(byLabel);
  var out = [];
  for (i = 0; i < labels.length; i++) {
    var lab = labels[i];
    var entry = byLabel[lab];
    var keyList = Object.keys(entry.keys);
    var qty = 0;
    // 別名對同一銷量：取最大值（非加總）
    for (k = 0; k < keyList.length; k++) {
      if (sold[keyList[k]] != null) {
        qty = Math.max(qty, sold[keyList[k]]);
      }
    }
    entry.qty = qty;
    entry.revenue = qty * (entry.price || 0);
    out.push(entry);
  }

  out.sort(function (a, b) { return b.revenue - a.revenue; });
  return out;
}

/* ========== 寫出 ========== */

function writePieTickets_(ss, rows) {
  var headers = [
    'row_id', 'channel', 'tally_column_hint', 'ticket_type', 'chart_label',
    'sales_qty', 'revenue', 'unit_price', 'updated_at'
  ];
  var body = [];
  var i;
  for (i = 0; i < rows.length; i++) {
    var x = rows[i];
    body.push([
      'TKT-' + (i + 1),
      'ticket',
      x.tally_column_hint || x.label,
      x.label,
      x.label, // pie 顯示名稱
      x.qty,
      x.revenue,
      x.price || 0,
      new Date()
    ]);
  }
  writeSheetFast_(ss, PIE_.OUT_TICKETS, headers, body, '#fce8e6');
}

function writePieMerch_(ss, rows) {
  var headers = [
    'row_id', 'channel', 'tally_column_hint', 'product_label', 'chart_label',
    'sales_qty', 'revenue', 'unit_price', 'updated_at'
  ];
  var body = [];
  var i;
  for (i = 0; i < rows.length; i++) {
    var x = rows[i];
    body.push([
      'MRC-' + (i + 1),
      'merch',
      x.tally_column_hint || x.label,
      x.label,
      x.label,
      x.qty,
      x.revenue,
      x.price || 0,
      new Date()
    ]);
  }
  writeSheetFast_(ss, PIE_.OUT_MERCH, headers, body, '#e6f4ea');
}

function writePieAll_(ss, tickets, merch) {
  var headers = [
    'row_id', 'channel', 'channel_zh', 'tally_column_hint', 'item_label', 'chart_label',
    'sales_qty', 'revenue', 'unit_price', 'updated_at'
  ];
  var body = [];
  var i;
  for (i = 0; i < tickets.length; i++) {
    var t = tickets[i];
    body.push([
      'ALL-T-' + (i + 1), 'ticket', '票務',
      t.tally_column_hint || t.label, t.label, t.label,
      t.qty, t.revenue, t.price || 0, new Date()
    ]);
  }
  for (i = 0; i < merch.length; i++) {
    var m = merch[i];
    body.push([
      'ALL-M-' + (i + 1), 'merch', '商品',
      m.tally_column_hint || m.label, m.label, m.label,
      m.qty, m.revenue, m.price || 0, new Date()
    ]);
  }
  writeSheetFast_(ss, PIE_.OUT_ALL, headers, body, '#fff2cc');
}

function writeSheetFast_(ss, name, headers, body, headerColor) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  var cols = headers.length;
  var rows = 1 + body.length;
  if (rows < 1) rows = 1;

  if (rows > sh.getMaxRows()) {
    sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  }
  if (cols > sh.getMaxColumns()) {
    sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
  }

  var oldLast = Math.min(Math.max(sh.getLastRow(), rows), 100);
  var oldCols = Math.min(Math.max(sh.getLastColumn(), cols), 15);
  if (oldLast >= 1 && oldCols >= 1) {
    try {
      rg_(sh, 1, 1, oldLast, oldCols).clearContent();
    } catch (e) {}
  }

  var all = [headers];
  var i;
  for (i = 0; i < body.length; i++) all.push(body[i]);
  rg_(sh, 1, 1, all.length, cols).setValues(all);

  try {
    rg_(sh, 1, 1, 1, cols).setFontWeight('bold').setBackground(headerColor);
    sh.setFrozenRows(1);
    // 金額格式
    var revCol = headers.indexOf('revenue') + 1;
    if (body.length && revCol > 0) {
      rg_(sh, 2, revCol, body.length, 1).setNumberFormat('"$"#,##0');
    }
  } catch (e2) {}
}

/* ========== utils ========== */

function mergeSold_(a, b) {
  var keys = Object.keys(b);
  var i;
  for (i = 0; i < keys.length; i++) {
    a[keys[i]] = (a[keys[i]] || 0) + b[keys[i]];
  }
}

function addSold_(sold, key, qty) {
  var k = normKey_(key);
  if (!k) return;
  sold[k] = (sold[k] || 0) + qty;
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
