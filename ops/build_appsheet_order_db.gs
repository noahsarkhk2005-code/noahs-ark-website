/**
 * AppSheet 訂單 + 圖表資料庫（適合讀取／Pie Chart）
 *
 * 來源：NA_Tickets / NA_Merch + Order_Categories（tally_column_hint、list_price）
 *
 * 產出（請全部加入 AppSheet）：
 *
 * 【訂單明細 — 一單一列】
 *   AS_TicketOrders   票務訂單
 *   AS_MerchOrders    商品訂單
 *
 * 【圖表專用 — 欄位少、給 Pie】
 *   AS_Chart_Tickets  label + sales_qty + revenue
 *   AS_Chart_Merch    label + sales_qty + revenue
 *
 * AppSheet Pie 設定（重要）：
 *   Chart columns 只勾：label + revenue（或 label + sales_qty）
 *   不要勾 _RowNumber
 *   主鍵：order_key / chart_key（Text）
 *
 * 執行：buildAppSheetOrderDb
 */

var DB_ = {
  TICKETS: 'NA_Tickets',
  MERCH: 'NA_Merch',
  CAT: 'Order_Categories',
  OUT_TICKET_ORDERS: 'AS_TicketOrders',
  OUT_MERCH_ORDERS: 'AS_MerchOrders',
  OUT_CHART_TICKETS: 'AS_Chart_Tickets',
  OUT_CHART_MERCH: 'AS_Chart_Merch'
};

function rg_(sh, r, c, nR, nC) {
  return sh.getRange(r, c, nR, nC);
}

function buildAppSheetOrderDb() {
  var ss = SpreadsheetApp.getActive();
  var cat = loadCatalog_(ss);
  var priceByHint = cat.priceByHint;
  var labelListTicket = cat.ticketLabels;
  var labelListMerch = cat.merchLabels;

  var ticketOrders = buildTicketOrderRows_(ss, priceByHint);
  var merchOrders = buildMerchOrderRows_(ss, priceByHint);

  SpreadsheetApp.flush();
  writeTable_(ss, DB_.OUT_TICKET_ORDERS, ticketOrderHeaders_(), ticketOrders, '#fce8e6');
  SpreadsheetApp.flush();
  writeTable_(ss, DB_.OUT_MERCH_ORDERS, merchOrderHeaders_(), merchOrders, '#e6f4ea');
  SpreadsheetApp.flush();

  var chartTickets = summarizeChart_(ticketOrders, 'ticket_label', labelListTicket);
  var chartMerch = summarizeChart_(merchOrders, 'product_label', labelListMerch);

  writeTable_(ss, DB_.OUT_CHART_TICKETS, chartHeaders_('ticket'), chartTickets, '#f4cccc');
  SpreadsheetApp.flush();
  writeTable_(ss, DB_.OUT_CHART_MERCH, chartHeaders_('merch'), chartMerch, '#d9ead3');
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    'AppSheet 資料庫已重建\n\n' +
    'AS_TicketOrders: ' + ticketOrders.length + ' 單\n' +
    'AS_MerchOrders: ' + merchOrders.length + ' 單\n' +
    'AS_Chart_Tickets: ' + chartTickets.length + ' 類（Pie）\n' +
    'AS_Chart_Merch: ' + chartMerch.length + ' 類（Pie）\n\n' +
    'Pie 請用 AS_Chart_* 表\n' +
    'Chart columns 只選：label + revenue（或 sales_qty）\n' +
    '不要選 _RowNumber'
  );
}

/* ========== 表頭（固定、給 AppSheet） ========== */

function ticketOrderHeaders_() {
  return [
    'order_key',        // KEY
    'submission_id',
    'submitted_at',
    'customer_name',
    'email',
    'phone',
    'metal_pass',
    'ticket_label',     // = tally_column_hint 顯示名
    'sales_qty',
    'unit_price',
    'revenue',
    'payment_proof',
    'channel',
    'updated_at'
  ];
}

function merchOrderHeaders_() {
  return [
    'order_key',
    'submission_id',
    'submitted_at',
    'customer_name',
    'email',
    'phone',
    'metal_pass',
    'product_label',    // = tally_column_hint 顯示名
    'sales_qty',
    'unit_price',
    'revenue',
    'payment_proof',
    'channel',
    'updated_at'
  ];
}

/** 圖表專用：欄位極簡，避免 AppSheet 亂勾 _RowNumber */
function chartHeaders_(kind) {
  // kind 僅註解用
  return [
    'chart_key',   // KEY
    'label',       // Pie 分類（tally_column_hint）
    'sales_qty',   // 總銷量 Number
    'revenue',     // 總銷售額 Number
    'channel'      // ticket | merch
  ];
}

/* ========== Catalog ========== */

function loadCatalog_(ss) {
  var priceByHint = {};
  var ticketLabels = [];
  var merchLabels = [];
  var seenT = {};
  var seenM = {};

  var sh = ss.getSheetByName(DB_.CAT);
  if (!sh || sh.getLastRow() < 2) {
    return { priceByHint: priceByHint, ticketLabels: ticketLabels, merchLabels: merchLabels };
  }

  var lastCol = Math.min(sh.getLastColumn(), 25);
  var lastRow = sh.getLastRow();
  var headers = rg_(sh, 1, 1, 1, lastCol).getValues()[0];
  var idx = headerIndexMap_(headers);
  var data = rg_(sh, 2, 1, lastRow - 1, lastCol).getValues();
  var r;
  for (r = 0; r < data.length; r++) {
    if (idx.level != null && Number(data[r][idx.level]) !== 3) continue;
    var channel = String(data[r][idx.channel] || '').toLowerCase();
    var hint = str_(data[r], idx, 'tally_column_hint');
    var form = str_(data[r], idx, 'form_option_label');
    var nameZh = str_(data[r], idx, 'name_zh');
    var label = hint || form || nameZh;
    if (!label) continue;
    var price = toMoney_(idx.list_price != null ? data[r][idx.list_price] : 0);
    if (!price) price = guessPrice_(label);

    // 多個別名對到同一 label 價錢
    setPrice_(priceByHint, label, price);
    setPrice_(priceByHint, hint, price);
    setPrice_(priceByHint, form, price);
    setPrice_(priceByHint, nameZh, price);
    setPrice_(priceByHint, stripEmoji_(label), price);

    if (channel === 'ticket' && !seenT[label]) {
      seenT[label] = true;
      ticketLabels.push(label);
    }
    if (channel === 'merch' && !seenM[label]) {
      seenM[label] = true;
      merchLabels.push(label);
    }
  }
  return { priceByHint: priceByHint, ticketLabels: ticketLabels, merchLabels: merchLabels };
}

function setPrice_(map, label, price) {
  if (!label || !price) return;
  map[normKey_(label)] = { label: String(label).trim(), price: price };
  map[normKey_(stripEmoji_(label))] = { label: String(label).trim(), price: price };
}

/* ========== 訂單明細 ========== */

function buildTicketOrderRows_(ss, priceByHint) {
  var sh = ss.getSheetByName(DB_.TICKETS);
  var out = [];
  if (!sh || sh.getLastRow() < 2) return out;

  var lastCol = Math.min(sh.getLastColumn(), 20);
  var lastRow = sh.getLastRow();
  var headers = rg_(sh, 1, 1, 1, lastCol).getValues()[0];
  var data = rg_(sh, 2, 1, lastRow - 1, lastCol).getValues();
  var h = headerIndexMap_(headers);

  var colSid = firstCol_(h, ['Submission ID', 'submission_id']);
  var colAt = firstCol_(h, ['Submitted at', '提交時間']);
  var colName = firstCol_(h, ['💀 Name/稱呼', 'Name/稱呼', 'Name', '稱呼']);
  var colEmail = firstCol_(h, ['📩 Email/電郵', 'Email/電郵', 'Email']);
  var colTel = firstCol_(h, ['☎️ Tel/電話號碼', 'Tel/電話號碼', 'Tel', '電話']);
  var colPass = firstCol_(h, ['Metal Pass 會員編號', 'Metal Pass']);
  var colType = firstCol_(h, ['門票類別', '票種']);
  var colQty = firstCol_(h, ['數量', 'Qty']);
  var colPrice = firstCol_(h, ['單價', 'Price']);
  var colProof = firstCol_(h, ['Payment Capture', '付款截圖', '請上傳付款截圖']);

  var r, c;
  var n = 0;
  var now = new Date();

  // 新格式：門票類別 + 數量
  if (colType != null && colQty != null) {
    for (r = 0; r < data.length; r++) {
      var sid = colSid != null ? String(data[r][colSid] || '').trim() : '';
      var type = String(data[r][colType] || '').trim();
      var qty = toQty_(data[r][colQty]);
      if (!type || qty <= 0) continue;
      if (!sid) sid = 'ROW' + (r + 2);

      var unit = colPrice != null ? toMoney_(data[r][colPrice]) : 0;
      var resolved = resolveLabelPrice_(type, priceByHint, unit);
      n++;
      out.push([
        'TORD-' + sid + '-' + n,
        sid,
        colAt != null ? data[r][colAt] : '',
        colName != null ? data[r][colName] : '',
        colEmail != null ? data[r][colEmail] : '',
        colTel != null ? data[r][colTel] : '',
        colPass != null ? data[r][colPass] : '',
        resolved.label,           // ticket_label = tally_column_hint 風格
        qty,
        resolved.price,
        qty * resolved.price,
        colProof != null ? data[r][colProof] : '',
        'ticket',
        now
      ]);
    }
    return out;
  }

  // 舊格式：各票種欄
  for (c = 0; c < headers.length; c++) {
    var hh = String(headers[c] || '').trim();
    if (!isTicketHeader_(hh)) continue;
    for (r = 0; r < data.length; r++) {
      var q = toQty_(data[r][c]);
      if (q <= 0) continue;
      var sid2 = colSid != null ? String(data[r][colSid] || '').trim() : ('ROW' + (r + 2));
      var res = resolveLabelPrice_(hh, priceByHint, 0);
      n++;
      out.push([
        'TORD-' + sid2 + '-' + n,
        sid2,
        colAt != null ? data[r][colAt] : '',
        colName != null ? data[r][colName] : '',
        colEmail != null ? data[r][colEmail] : '',
        colTel != null ? data[r][colTel] : '',
        colPass != null ? data[r][colPass] : '',
        res.label,
        q,
        res.price,
        q * res.price,
        colProof != null ? data[r][colProof] : '',
        'ticket',
        now
      ]);
    }
  }
  return out;
}

function buildMerchOrderRows_(ss, priceByHint) {
  var sh = ss.getSheetByName(DB_.MERCH);
  var out = [];
  if (!sh || sh.getLastRow() < 2) return out;

  var lastCol = Math.min(sh.getLastColumn(), 25);
  var lastRow = sh.getLastRow();
  var headers = rg_(sh, 1, 1, 1, lastCol).getValues()[0];
  var data = rg_(sh, 2, 1, lastRow - 1, lastCol).getValues();
  var h = headerIndexMap_(headers);

  var colSid = firstCol_(h, ['Submission ID', 'submission_id']);
  var colAt = firstCol_(h, ['Submitted at', '提交時間']);
  var colName = firstCol_(h, ['💀 Name/稱呼', 'Name/稱呼', 'Name', '稱呼']);
  var colEmail = firstCol_(h, ['📩 Email/電郵', 'Email/電郵', 'Email']);
  var colTel = firstCol_(h, ['☎️ Tel/電話號碼', 'Tel/電話號碼', 'Tel', '電話']);
  var colPass = firstCol_(h, ['Metal Pass 會員編號', 'Metal Pass']);
  var colCat = firstCol_(h, ['商品分欄', '商品類別']);
  var colQty = firstCol_(h, ['數量', 'Qty']);
  var colPrice = firstCol_(h, ['單價', 'Price']);
  var colProof = firstCol_(h, ['請上傳付款截圖', 'Payment Capture', '付款截圖']);

  var r, c;
  var n = 0;
  var now = new Date();

  // 新格式
  if (colCat != null && colQty != null) {
    for (r = 0; r < data.length; r++) {
      var sid = colSid != null ? String(data[r][colSid] || '').trim() : '';
      var cat = String(data[r][colCat] || '').trim();
      var qty = toQty_(data[r][colQty]);
      if (!cat || qty <= 0) continue;
      if (!sid) sid = 'ROW' + (r + 2);
      var unit = colPrice != null ? toMoney_(data[r][colPrice]) : 0;
      var res = resolveLabelPrice_(cat, priceByHint, unit);
      n++;
      out.push([
        'MORD-' + sid + '-' + n,
        sid,
        colAt != null ? data[r][colAt] : '',
        colName != null ? data[r][colName] : '',
        colEmail != null ? data[r][colEmail] : '',
        colTel != null ? data[r][colTel] : '',
        colPass != null ? data[r][colPass] : '',
        res.label,
        qty,
        res.price,
        qty * res.price,
        colProof != null ? data[r][colProof] : '',
        'merch',
        now
      ]);
    }
  }

  // 舊商品欄（欄名常 = tally_column_hint）
  for (c = 0; c < headers.length; c++) {
    var hh = String(headers[c] || '').trim();
    if (!isMerchProductHeader_(hh)) continue;
    if (c === colCat || c === colQty || c === colPrice) continue;
    for (r = 0; r < data.length; r++) {
      var q = toQty_(data[r][c]);
      if (q <= 0) continue;
      var sid2 = colSid != null ? String(data[r][colSid] || '').trim() : ('ROW' + (r + 2));
      var res2 = resolveLabelPrice_(hh, priceByHint, 0);
      n++;
      out.push([
        'MORD-' + sid2 + '-' + n,
        sid2,
        colAt != null ? data[r][colAt] : '',
        colName != null ? data[r][colName] : '',
        colEmail != null ? data[r][colEmail] : '',
        colTel != null ? data[r][colTel] : '',
        colPass != null ? data[r][colPass] : '',
        res2.label,
        q,
        res2.price,
        q * res2.price,
        colProof != null ? data[r][colProof] : '',
        'merch',
        now
      ]);
    }
  }
  return out;
}

/**
 * 把 Tally 原始字串對到 catalog 的 tally_column_hint 顯示名 + 價錢
 */
function resolveLabelPrice_(raw, priceByHint, unitFromRow) {
  var rawS = String(raw || '').trim();
  var k = normKey_(rawS);
  var k2 = normKey_(stripEmoji_(rawS));
  var hit = priceByHint[k] || priceByHint[k2];
  if (hit) {
    return {
      label: hit.label, // 使用 catalog 的 tally_column_hint 原文
      price: unitFromRow > 0 ? unitFromRow : hit.price
    };
  }
  var price = unitFromRow > 0 ? unitFromRow : guessPrice_(rawS);
  return { label: rawS, price: price };
}

/* ========== 圖表彙總 ========== */

function summarizeChart_(orderRows, labelIndexName, allLabels) {
  // orderRows 是二維陣列；找 label / qty / revenue 欄位 index
  // ticket: label=7, qty=8, revenue=10
  // merch:  label=7, qty=8, revenue=10
  var LABEL = 7;
  var QTY = 8;
  var REV = 10;

  var map = {};
  var i;
  // 先放入 catalog 全部 label（即使 0 銷量，方便完整圖例；pie 可只顯示 >0）
  if (allLabels && allLabels.length) {
    for (i = 0; i < allLabels.length; i++) {
      map[allLabels[i]] = { label: allLabels[i], qty: 0, revenue: 0 };
    }
  }

  for (i = 0; i < orderRows.length; i++) {
    var row = orderRows[i];
    var lab = String(row[LABEL] || '').trim();
    if (!lab) continue;
    if (!map[lab]) map[lab] = { label: lab, qty: 0, revenue: 0 };
    map[lab].qty += toQty_(row[QTY]);
    map[lab].revenue += toMoney_(row[REV]);
  }

  var keys = Object.keys(map);
  keys.sort(function (a, b) { return map[b].revenue - map[a].revenue; });

  var out = [];
  var channel = labelIndexName === 'ticket_label' ? 'ticket' : 'merch';
  for (i = 0; i < keys.length; i++) {
    var x = map[keys[i]];
    // 圖表建議只顯示有銷量的；若要顯示 0 也可保留
    if (x.qty <= 0 && x.revenue <= 0) continue;
    out.push([
      (channel === 'ticket' ? 'CT-' : 'CM-') + (i + 1),
      x.label,       // label = tally_column_hint
      x.qty,         // sales_qty  Number
      x.revenue,     // revenue    Number
      channel
    ]);
  }
  return out;
}

/* ========== 寫表 ========== */

function writeTable_(ss, name, headers, body, headerColor) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  var cols = headers.length;
  var nBody = body.length;
  var totalRows = 1 + Math.max(nBody, 0);

  if (totalRows > sh.getMaxRows()) {
    sh.insertRowsAfter(sh.getMaxRows(), totalRows - sh.getMaxRows() + 5);
  }
  if (cols > sh.getMaxColumns()) {
    sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns() + 2);
  }

  var clearR = Math.min(Math.max(sh.getLastRow(), totalRows), 500);
  var clearC = Math.min(Math.max(sh.getLastColumn(), cols), 20);
  try {
    if (clearR >= 1 && clearC >= 1) rg_(sh, 1, 1, clearR, clearC).clearContent();
  } catch (e) {}

  var all = [headers];
  var i;
  for (i = 0; i < nBody; i++) all.push(body[i]);
  rg_(sh, 1, 1, all.length, cols).setValues(all);

  try {
    rg_(sh, 1, 1, 1, cols).setFontWeight('bold').setBackground(headerColor || '#eee');
    sh.setFrozenRows(1);
    // Number formats for chart reliability
    var qtyCol = headers.indexOf('sales_qty') + 1;
    var revCol = headers.indexOf('revenue') + 1;
    if (nBody > 0 && qtyCol > 0) rg_(sh, 2, qtyCol, nBody, 1).setNumberFormat('0');
    if (nBody > 0 && revCol > 0) rg_(sh, 2, revCol, nBody, 1).setNumberFormat('0');
  } catch (e2) {}
}

/* ========== utils ========== */

function str_(row, idx, name) {
  if (!idx || idx[name] == null) return '';
  return String(row[idx[name]] || '').trim();
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

function guessPrice_(h) {
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
