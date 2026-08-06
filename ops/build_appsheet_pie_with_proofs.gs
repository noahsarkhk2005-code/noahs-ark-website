/**
 * AS_Pie_Tickets / AS_Pie_Merch
 * = 每張 Tally 訂單一列 + 付款截圖 URL
 *
 * 分類名稱 chart_label = Order_Categories.tally_column_hint
 * sales_qty / unit_price / revenue 正確數字格式
 * payment_proof_url = 該訂單截圖
 *
 * 執行：buildPieOrdersWithProofs
 *
 * AppSheet：
 *   - 主鍵 row_id
 *   - payment_proof_url 類型 Image 或 URL
 *   - Pie 可用 AS_Chart_* 彙總表，或對本表 Group by chart_label
 */

var PIEP_ = {
  TICKETS: 'NA_Tickets',
  MERCH: 'NA_Merch',
  CAT: 'Order_Categories',
  OUT_T: 'AS_Pie_Tickets',
  OUT_M: 'AS_Pie_Merch',
  OUT_CT: 'AS_Chart_Tickets',
  OUT_CM: 'AS_Chart_Merch'
};

function rg_(sh, r, c, nR, nC) {
  return sh.getRange(r, c, nR, nC);
}

function buildPieOrdersWithProofs() {
  var ss = SpreadsheetApp.getActive();
  var catalog = loadCatalogFull_(ss);

  var ticketOrders = buildTicketOrderDetail_(ss, catalog);
  var merchOrders = buildMerchOrderDetail_(ss, catalog);

  SpreadsheetApp.flush();
  writePieTicketOrders_(ss, ticketOrders);
  SpreadsheetApp.flush();
  writePieMerchOrders_(ss, merchOrders);
  SpreadsheetApp.flush();

  // 同時更新彙總 Chart 表（Pie 用，無截圖）
  writeChartFromOrders_(ss, PIEP_.OUT_CT, ticketOrders, 'ticket');
  SpreadsheetApp.flush();
  writeChartFromOrders_(ss, PIEP_.OUT_CM, merchOrders, 'merch');

  SpreadsheetApp.getUi().alert(
    '已更新（每單一列 + 截圖 URL）\n\n' +
    'AS_Pie_Tickets: ' + ticketOrders.length + ' 單\n' +
    'AS_Pie_Merch: ' + merchOrders.length + ' 單\n' +
    'AS_Chart_Tickets / Merch: 彙總供 Pie\n\n' +
    '欄位 payment_proof_url = 訂單截圖\n' +
    'AppSheet Sync 後，Image 欄可顯示截圖'
  );
}

/* ========== Catalog ========== */

function loadCatalogFull_(ss) {
  // returns { byKey: norm-> {label, price, channel}, ticketLabels:[], merchLabels:[] }
  var byKey = {};
  var sh = ss.getSheetByName(PIEP_.CAT);
  if (!sh || sh.getLastRow() < 2) {
    seedFallbackCatalog_(byKey);
    return { byKey: byKey };
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
    var hint = cell_(data[r], idx, 'tally_column_hint');
    var form = cell_(data[r], idx, 'form_option_label');
    var nameZh = cell_(data[r], idx, 'name_zh');
    var label = hint || form || nameZh;
    if (!label) continue;
    var price = toMoney_(idx.list_price != null ? data[r][idx.list_price] : 0);
    if (!price) price = guessPrice_(label);
    registerLabel_(byKey, label, price, channel);
    registerLabel_(byKey, hint, price, channel);
    registerLabel_(byKey, form, price, channel);
    registerLabel_(byKey, nameZh, price, channel);
    registerLabel_(byKey, stripEmoji_(label), price, channel);
  }
  return { byKey: byKey };
}

function seedFallbackCatalog_(byKey) {
  registerLabel_(byKey, '會員特級優惠 HKD300', 300, 'ticket');
  registerLabel_(byKey, '🎫 早鳥門票 HKD300', 300, 'ticket');
  registerLabel_(byKey, '🎫 預售 HKD350', 350, 'ticket');
  registerLabel_(byKey, 'Ark T-Shirt HKD280', 280, 'merch');
  registerLabel_(byKey, 'Ark Tower/毛巾 HKD80', 80, 'merch');
  registerLabel_(byKey, 'Ark Keychain/鎖匙扣 HKD120', 120, 'merch');
  registerLabel_(byKey, 'Ark BigPack HKD200', 200, 'merch');
  registerLabel_(byKey, 'Ark TinyPack HKD60', 60, 'merch');
  registerLabel_(byKey, 'Ark Tote Bag HKD120', 120, 'merch');
  registerLabel_(byKey, '服飾', 280, 'merch');
  registerLabel_(byKey, '毛巾', 80, 'merch');
  registerLabel_(byKey, '配件', 120, 'merch');
  registerLabel_(byKey, '套裝 Pack', 200, 'merch');
  registerLabel_(byKey, '袋類', 120, 'merch');
}

function registerLabel_(byKey, label, price, channel) {
  if (!label) return;
  var L = String(label).trim();
  var k = normKey_(L);
  if (!k) return;
  if (!byKey[k] || L.length >= byKey[k].label.length) {
    byKey[k] = { label: L, price: price, channel: channel || '' };
  }
}

/**
 * 永遠回傳 { label, price, channel }，不會是 null
 */
function resolveLabel_(raw, catalog, preferChannel) {
  var s = String(raw || '').trim();
  var fallback = {
    label: s || '(未分類)',
    price: guessPrice_(s),
    channel: preferChannel || ''
  };
  if (!s) return fallback;
  if (!catalog || !catalog.byKey) return fallback;

  var k = normKey_(s);
  var k2 = normKey_(stripEmoji_(s));
  var hit = catalog.byKey[k] || catalog.byKey[k2];
  if (hit && hit.label) return hit;

  var keys = Object.keys(catalog.byKey);
  var i;
  for (i = 0; i < keys.length; i++) {
    var ck = keys[i];
    var item = catalog.byKey[ck];
    if (!item || !item.label) continue;
    if (preferChannel && item.channel && item.channel !== preferChannel) continue;
    if (k && ck && (k.indexOf(ck) >= 0 || ck.indexOf(k) >= 0 || k2.indexOf(ck) >= 0 || ck.indexOf(k2) >= 0)) {
      return item;
    }
  }

  // 關鍵字（findByRegex 可能找不到 → 必須 fallback）
  var reHit = null;
  if (/早鳥|early/i.test(s)) reHit = findByRegex_(catalog, /早鳥|early/i, preferChannel);
  else if (/預售|advanced|presale/i.test(s)) reHit = findByRegex_(catalog, /預售|advanced|presale/i, preferChannel);
  else if (/會員|metal\s*pass|特級/i.test(s)) reHit = findByRegex_(catalog, /會員|metal|特級/i, preferChannel);
  else if (/tee|t-shirt|服飾/i.test(s)) reHit = findByRegex_(catalog, /tee|t-shirt|服飾|T-Shirt/i, preferChannel);
  else if (/towel|毛巾|tower/i.test(s)) reHit = findByRegex_(catalog, /towel|毛巾|tower/i, preferChannel);
  else if (/keychain|鎖匙/i.test(s)) reHit = findByRegex_(catalog, /keychain|鎖匙/i, preferChannel);
  else if (/配件/i.test(s)) reHit = findByRegex_(catalog, /配件|keychain|鎖匙|patch|布章/i, preferChannel);
  else if (/tote|袋/i.test(s)) reHit = findByRegex_(catalog, /tote|袋/i, preferChannel);
  else if (/bigpack|big pack/i.test(s)) reHit = findByRegex_(catalog, /bigpack|BigPack/i, preferChannel);
  else if (/tinypack|tiny pack/i.test(s)) reHit = findByRegex_(catalog, /tinypack|TinyPack/i, preferChannel);
  else if (/pack/i.test(s)) reHit = findByRegex_(catalog, /pack|Pack/i, preferChannel);

  if (reHit && reHit.label) return reHit;
  return fallback;
}

function findByRegex_(catalog, re, preferChannel) {
  if (!catalog || !catalog.byKey) return null;
  var keys = Object.keys(catalog.byKey);
  var i;
  for (i = 0; i < keys.length; i++) {
    var item = catalog.byKey[keys[i]];
    if (!item || !item.label) continue;
    if (preferChannel && item.channel && item.channel !== preferChannel) continue;
    if (re.test(item.label)) return item;
  }
  return null;
}

/* ========== 票務訂單明細（含截圖） ========== */

function buildTicketOrderDetail_(ss, catalog) {
  var sh = ss.getSheetByName(PIEP_.TICKETS);
  var out = [];
  if (!sh || sh.getLastRow() < 2) return out;

  var lastCol = Math.min(sh.getLastColumn(), 16);
  var lastRow = sh.getLastRow();
  var headers = rg_(sh, 1, 1, 1, lastCol).getValues()[0];
  var data = rg_(sh, 2, 1, lastRow - 1, lastCol).getValues();
  var h = headerIndexMap_(headers);

  var cSid = firstCol_(h, ['Submission ID']);
  var cAt = firstCol_(h, ['Submitted at']);
  var cName = firstCol_(h, ['💀 Name/稱呼', 'Name/稱呼', 'Name', '稱呼']);
  var cEmail = firstCol_(h, ['📩 Email/電郵', 'Email']);
  var cTel = firstCol_(h, ['☎️ Tel/電話號碼', 'Tel', '電話']);
  var cPass = firstCol_(h, ['Metal Pass 會員編號', 'Metal Pass']);
  var cType = firstCol_(h, ['門票類別', '票種']);
  var cQty = firstCol_(h, ['數量', 'Qty']);
  var cPrice = firstCol_(h, ['單價']);
  var cProof = firstCol_(h, ['Payment Capture', '付款截圖', '請上傳付款截圖']);
  var cTotal = firstCol_(h, ['Total Amount', '總金額']);

  var now = new Date();
  var r, c;
  var seq = 0;

  if (cType != null && cQty != null) {
    for (r = 0; r < data.length; r++) {
      var type = String(data[r][cType] || '').trim();
      var qty = toQty_(data[r][cQty]);
      if (!type || qty <= 0) continue;
      var sid = cSid != null ? String(data[r][cSid] || '').trim() : '';
      if (!sid) sid = 'R' + (r + 2);
      var hit = resolveLabel_(type, catalog, 'ticket') || { label: type, price: guessPrice_(type) };
      var unit = cPrice != null ? toMoney_(data[r][cPrice]) : 0;
      if (!unit || unit > 10000) unit = hit.price || 0; // 避免日期序號當單價
      if (!unit) unit = guessPrice_(type);
      var proof = normalizeProofUrl_(cProof != null ? data[r][cProof] : '');
      seq++;
      out.push(makeOrderObj_({
        row_id: 'TKT-' + sid + '-' + seq,
        submission_id: sid,
        submitted_at: cAt != null ? data[r][cAt] : '',
        customer_name: cName != null ? data[r][cName] : '',
        email: cEmail != null ? data[r][cEmail] : '',
        phone: cTel != null ? data[r][cTel] : '',
        metal_pass: cPass != null ? data[r][cPass] : '',
        chart_label: hit.label || type,
        tally_column_hint: hit.label || type,
        sales_qty: qty,
        unit_price: unit,
        revenue: qty * unit,
        payment_proof_url: proof,
        total_amount: cTotal != null ? toMoney_(data[r][cTotal]) : qty * unit,
        channel: 'ticket',
        updated_at: now
      }));
    }
    return out;
  }

  // 舊多欄
  for (c = 0; c < headers.length; c++) {
    var hh = String(headers[c] || '').trim();
    if (!isTicketHeader_(hh)) continue;
    for (r = 0; r < data.length; r++) {
      var q = toQty_(data[r][c]);
      if (q <= 0) continue;
      var sid2 = cSid != null ? String(data[r][cSid] || '').trim() : ('R' + (r + 2));
      var hit2 = resolveLabel_(hh, catalog, 'ticket') || { label: hh, price: guessPrice_(hh) };
      var proof2 = normalizeProofUrl_(cProof != null ? data[r][cProof] : '');
      var p2 = hit2.price || guessPrice_(hh) || 0;
      seq++;
      out.push(makeOrderObj_({
        row_id: 'TKT-' + sid2 + '-' + seq,
        submission_id: sid2,
        submitted_at: cAt != null ? data[r][cAt] : '',
        customer_name: cName != null ? data[r][cName] : '',
        email: cEmail != null ? data[r][cEmail] : '',
        phone: cTel != null ? data[r][cTel] : '',
        metal_pass: cPass != null ? data[r][cPass] : '',
        chart_label: hit2.label || hh,
        tally_column_hint: hit2.label || hh,
        sales_qty: q,
        unit_price: p2,
        revenue: q * p2,
        payment_proof_url: proof2,
        total_amount: cTotal != null ? toMoney_(data[r][cTotal]) : 0,
        channel: 'ticket',
        updated_at: now
      }));
    }
  }
  return out;
}

/* ========== 商品訂單明細（含截圖） ========== */

function buildMerchOrderDetail_(ss, catalog) {
  var sh = ss.getSheetByName(PIEP_.MERCH);
  var out = [];
  if (!sh || sh.getLastRow() < 2) return out;

  var lastCol = Math.min(sh.getLastColumn(), 25);
  var lastRow = sh.getLastRow();
  var headers = rg_(sh, 1, 1, 1, lastCol).getValues()[0];
  var data = rg_(sh, 2, 1, lastRow - 1, lastCol).getValues();
  var h = headerIndexMap_(headers);

  var cSid = firstCol_(h, ['Submission ID']);
  var cAt = firstCol_(h, ['Submitted at']);
  var cName = firstCol_(h, ['💀 Name/稱呼', 'Name/稱呼', 'Name', '稱呼']);
  var cEmail = firstCol_(h, ['📩 Email/電郵', 'Email']);
  var cTel = firstCol_(h, ['☎️ Tel/電話號碼', 'Tel', '電話']);
  var cPass = firstCol_(h, ['Metal Pass 會員編號', 'Metal Pass']);
  var cCat = firstCol_(h, ['商品分欄', '商品類別']);
  var cQty = firstCol_(h, ['數量', 'Qty']);
  var cPrice = firstCol_(h, ['單價']);
  var cProof = firstCol_(h, ['請上傳付款截圖', 'Payment Capture', '付款截圖']);
  var cTotal = firstCol_(h, ['Total Amount', '總金額']);

  var now = new Date();
  var r, c;
  var seq = 0;
  var seen = {}; // sid+label+qty 防重複（新舊格式並存）

  // 新格式 IJK
  if (cCat != null && cQty != null) {
    for (r = 0; r < data.length; r++) {
      var cat = String(data[r][cCat] || '').trim();
      var qty = toQty_(data[r][cQty]);
      if (!cat || qty <= 0) continue;
      var sid = cSid != null ? String(data[r][cSid] || '').trim() : '';
      if (!sid) sid = 'R' + (r + 2);
      var hit = resolveLabel_(cat, catalog, 'merch') || { label: cat, price: guessPrice_(cat) };
      var unit = cPrice != null ? toMoney_(data[r][cPrice]) : 0;
      if (!unit || unit > 10000) unit = hit.price || 0;
      if (!unit) unit = guessPrice_(cat);
      var lab = hit.label || cat || '(未分類)';
      var proof = normalizeProofUrl_(cProof != null ? data[r][cProof] : '');
      var uk = sid + '|' + lab + '|' + qty + '|ijk';
      if (seen[uk]) continue;
      seen[uk] = true;
      seq++;
      out.push(makeOrderObj_({
        row_id: 'MRC-' + sid + '-' + seq,
        submission_id: sid,
        submitted_at: cAt != null ? data[r][cAt] : '',
        customer_name: cName != null ? data[r][cName] : '',
        email: cEmail != null ? data[r][cEmail] : '',
        phone: cTel != null ? data[r][cTel] : '',
        metal_pass: cPass != null ? data[r][cPass] : '',
        chart_label: lab,
        tally_column_hint: lab,
        sales_qty: qty,
        unit_price: unit,
        revenue: qty * unit,
        payment_proof_url: proof,
        total_amount: cTotal != null ? toMoney_(data[r][cTotal]) : qty * unit,
        channel: 'merch',
        updated_at: now
      }));
    }
  }

  // 舊商品數量欄（欄名 = tally_column_hint）
  for (c = 0; c < headers.length; c++) {
    var hh = String(headers[c] || '').trim();
    if (!isMerchProductHeader_(hh)) continue;
    if (c === cCat || c === cQty || c === cPrice) continue;
    for (r = 0; r < data.length; r++) {
      var q = toQty_(data[r][c]);
      if (q <= 0) continue;
      var sid2 = cSid != null ? String(data[r][cSid] || '').trim() : ('R' + (r + 2));
      var hit2 = resolveLabel_(hh, catalog, 'merch') || { label: hh, price: guessPrice_(hh) };
      var lab2 = hit2.label || hh || '(未分類)';
      var p2 = hit2.price || guessPrice_(hh) || 0;
      var proof2 = normalizeProofUrl_(cProof != null ? data[r][cProof] : '');
      var uk2 = sid2 + '|' + lab2 + '|' + q + '|col';
      // 若同一 sid 已用 IJK 寫過同 label，跳過舊欄避免重複
      var ukIjk = sid2 + '|' + lab2 + '|' + q + '|ijk';
      if (seen[ukIjk] || seen[uk2]) continue;
      seen[uk2] = true;
      seq++;
      out.push(makeOrderObj_({
        row_id: 'MRC-' + sid2 + '-' + seq,
        submission_id: sid2,
        submitted_at: cAt != null ? data[r][cAt] : '',
        customer_name: cName != null ? data[r][cName] : '',
        email: cEmail != null ? data[r][cEmail] : '',
        phone: cTel != null ? data[r][cTel] : '',
        metal_pass: cPass != null ? data[r][cPass] : '',
        chart_label: lab2,
        tally_column_hint: lab2,
        sales_qty: q,
        unit_price: p2,
        revenue: q * p2,
        payment_proof_url: proof2,
        total_amount: cTotal != null ? toMoney_(data[r][cTotal]) : 0,
        channel: 'merch',
        updated_at: now
      }));
    }
  }
  return out;
}

function makeOrderObj_(o) {
  return o;
}

/* ========== 寫 AS_Pie_*（每單一列 + 截圖） ========== */

function writePieTicketOrders_(ss, orders) {
  var headers = [
    'row_id',
    'submission_id',
    'submitted_at',
    'customer_name',
    'email',
    'phone',
    'metal_pass',
    'tally_column_hint',
    'chart_label',
    'sales_qty',
    'unit_price',
    'revenue',
    'payment_proof_url',
    'total_amount',
    'channel',
    'updated_at'
  ];
  var body = [];
  var i;
  for (i = 0; i < orders.length; i++) {
    var o = orders[i];
    body.push([
      o.row_id,
      o.submission_id,
      o.submitted_at,
      o.customer_name,
      o.email,
      o.phone,
      o.metal_pass,
      o.tally_column_hint,
      o.chart_label,
      o.sales_qty,
      o.unit_price,
      o.revenue,
      o.payment_proof_url,
      o.total_amount,
      o.channel,
      o.updated_at
    ]);
  }
  writeTableNum_(ss, PIEP_.OUT_T, headers, body, '#fce8e6', {
    qty: 10, price: 11, rev: 12, dateCols: [3, 16]
  });
}

function writePieMerchOrders_(ss, orders) {
  var headers = [
    'row_id',
    'submission_id',
    'submitted_at',
    'customer_name',
    'email',
    'phone',
    'metal_pass',
    'tally_column_hint',
    'chart_label',
    'sales_qty',
    'unit_price',
    'revenue',
    'payment_proof_url',
    'total_amount',
    'channel',
    'updated_at'
  ];
  var body = [];
  var i;
  for (i = 0; i < orders.length; i++) {
    var o = orders[i];
    body.push([
      o.row_id,
      o.submission_id,
      o.submitted_at,
      o.customer_name,
      o.email,
      o.phone,
      o.metal_pass,
      o.tally_column_hint,
      o.chart_label,
      o.sales_qty,
      o.unit_price,
      o.revenue,
      o.payment_proof_url,
      o.total_amount,
      o.channel,
      o.updated_at
    ]);
  }
  writeTableNum_(ss, PIEP_.OUT_M, headers, body, '#e6f4ea', {
    qty: 10, price: 11, rev: 12, dateCols: [3, 16]
  });
}

/** 彙總表給 Pie（無截圖） */
function writeChartFromOrders_(ss, sheetName, orders, channel) {
  var map = {};
  var i;
  for (i = 0; i < orders.length; i++) {
    var o = orders[i];
    var lab = o.chart_label;
    if (!map[lab]) map[lab] = { label: lab, qty: 0, revenue: 0 };
    map[lab].qty += o.sales_qty;
    map[lab].revenue += o.revenue;
  }
  var keys = Object.keys(map);
  keys.sort(function (a, b) { return map[b].revenue - map[a].revenue; });
  var headers = ['chart_key', 'label', 'sales_qty', 'revenue', 'channel'];
  var body = [];
  for (i = 0; i < keys.length; i++) {
    var x = map[keys[i]];
    if (x.qty <= 0) continue;
    body.push([
      (channel === 'ticket' ? 'CT-' : 'CM-') + (i + 1),
      x.label,
      x.qty,
      x.revenue,
      channel
    ]);
  }
  writeTableNum_(ss, sheetName, headers, body, channel === 'ticket' ? '#f4cccc' : '#d9ead3', {
    qty: 3, rev: 4, price: 0, dateCols: []
  });
}

function writeTableNum_(ss, name, headers, body, color, fmt) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  var cols = headers.length;
  var nBody = body.length;
  var total = 1 + Math.max(nBody, 0);

  if (total + 2 > sh.getMaxRows()) {
    sh.insertRowsAfter(sh.getMaxRows(), total - sh.getMaxRows() + 5);
  }
  if (cols + 1 > sh.getMaxColumns()) {
    sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns() + 3);
  }

  var clearR = Math.min(Math.max(sh.getLastRow(), total), 800);
  var clearC = Math.min(Math.max(sh.getLastColumn(), cols), 25);
  try {
    rg_(sh, 1, 1, clearR, clearC).clearContent();
    rg_(sh, 1, 1, clearR, clearC).setNumberFormat('General');
  } catch (e) {}

  var all = [headers];
  var i;
  for (i = 0; i < nBody; i++) all.push(body[i]);
  rg_(sh, 1, 1, all.length, cols).setValues(all);

  rg_(sh, 1, 1, 1, cols).setFontWeight('bold').setBackground(color || '#eee');
  sh.setFrozenRows(1);

  if (nBody > 0 && fmt) {
    if (fmt.qty > 0) rg_(sh, 2, fmt.qty, nBody, 1).setNumberFormat('0');
    if (fmt.price > 0) rg_(sh, 2, fmt.price, nBody, 1).setNumberFormat('0');
    if (fmt.rev > 0) rg_(sh, 2, fmt.rev, nBody, 1).setNumberFormat('0');
    if (fmt.dateCols && fmt.dateCols.length) {
      var d;
      for (d = 0; d < fmt.dateCols.length; d++) {
        var dc = fmt.dateCols[d];
        if (dc > 0) {
          try {
            rg_(sh, 2, dc, nBody, 1).setNumberFormat('yyyy-mm-dd hh:mm');
          } catch (e2) {}
        }
      }
    }
  }
}

/**
 * 正規化截圖 URL
 * Tally 可能給字串 URL，或 Google Drive 連結
 */
function normalizeProofUrl_(v) {
  if (v === '' || v == null) return '';
  var s = String(v).trim();
  if (!s) return '';
  // 多個 URL 取第一個
  var m = s.match(/https?:\/\/[^\s,;"']+/);
  if (m) return m[0];
  return s;
}

/* ========== utils ========== */

function cell_(row, idx, name) {
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
  if (typeof v === 'number') {
    if (!isFinite(v)) return 0;
    // 排除日期序號誤判（>10000 不太像票價）
    if (v > 10000) return 0;
    return v;
  }
  var n = Number(String(v).replace(/[^0-9.\-]/g, ''));
  if (!isFinite(n) || n > 10000) return 0;
  return n;
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
