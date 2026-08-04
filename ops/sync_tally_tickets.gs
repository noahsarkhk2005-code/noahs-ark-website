/**
 * Noah's Ark — 同步 Tally 購票 → Orders
 * 貼到 Apps Script 後儲存，重新整理試算表
 * 選單：NOAHSARK → 同步 Tally 購票 → Orders
 *
 * Settings 請設：TICKET_SOURCE_TAB_NAME = 你的 Tally 分頁名（例：NA_Tickets）
 */

var TICKET_CFG_ = {
  SETTINGS_SHEET: 'Settings',
  ORDERS_SHEET: 'Orders',
  DEFAULT_SOURCE_TAB: 'NA_Tickets',
  SYNC_LOG_SHEET: '_TallyTicketSyncLog',
  EARLY_BIRD_PRICE: 300,
  ADVANCED_PRICE: 350,
  SOURCE: 'ticket_form',
  CHANNEL: 'tally'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('NOAHSARK')
    .addItem('同步 Tally 購票 → Orders', 'syncTallyTicketsToOrders')
    .addToUi();
}

function syncTallyTicketsToOrders() {
  var ss = SpreadsheetApp.getActive();
  var settings = readSettingsMap_(ss);
  var sourceName = settings.TICKET_SOURCE_TAB_NAME || TICKET_CFG_.DEFAULT_SOURCE_TAB;
  var source = ss.getSheetByName(sourceName);
  if (!source) {
    throw new Error('找不到來源分頁: ' + sourceName + '。請在 Settings 設 TICKET_SOURCE_TAB_NAME');
  }

  var orders = ss.getSheetByName(TICKET_CFG_.ORDERS_SHEET);
  if (!orders) {
    throw new Error('找不到 Orders 分頁');
  }

  var logSheet = ensureSyncLog_(ss);
  var done = loadSyncedIds_(logSheet);

  var data = source.getDataRange().getValues();
  if (!data || data.length < 2) {
    SpreadsheetApp.getUi().alert('來源分頁沒有資料列: ' + sourceName);
    return;
  }

  var headerRowIdx = 0;
  var r;
  for (r = 0; r < Math.min(5, data.length); r++) {
    if (rowHasAny_(data[r])) {
      headerRowIdx = r;
      break;
    }
  }

  var header = [];
  var c;
  for (c = 0; c < data[headerRowIdx].length; c++) {
    header.push(String(data[headerRowIdx][c] || '').trim());
  }

  var sidCol = findCol_(header, [
    'Submission ID', 'submission id', 'submission_id', 'SubmissionID',
    '提交 ID', '提交ID', '提交編號', 'Response ID'
  ]);
  if (sidCol < 0 && looksLikeSubmissionIdColumn_(data, headerRowIdx, 0)) {
    sidCol = 0;
  }
  if (sidCol < 0) {
    throw new Error(
      '來源分頁找不到 Submission ID。\n實際表頭:\n' +
      nonEmptyJoin_(header) +
      '\n\n請確認 TICKET_SOURCE_TAB_NAME = ' + sourceName
    );
  }

  var submittedCol = findCol_(header, [
    'Submitted at', 'submitted_at', 'Submitted At', '提交時間', 'Created at', 'Timestamp'
  ]);
  var totalCol = findCol_(header, [
    'Total Amount', 'total_amount', 'Total', '總金額', '總價', 'Amount'
  ]);
  var nameCol = findCol_(header, [
    'Name/稱呼', 'Name', '稱呼', '姓名', 'name'
  ]);
  var emailCol = findCol_(header, [
    'Email/電郵', 'Email', '電郵', 'email', 'E-mail'
  ]);
  var telCol = findCol_(header, [
    'Tel/電話號碼', 'Tel', '電話', '電話號碼', 'Phone', 'phone'
  ]);
  var passCol = findCol_(header, [
    'Metal Pass 會員編號', 'Metal Pass', 'METAL-PASS', '會員編號'
  ]);
  var ebCol = findCol_(header, [
    '早鳥門票 HKD300', '早鳥門票', '早鳥', 'Early Bird', 'EarlyBird'
  ]);
  var advCol = findCol_(header, [
    '預售 HKD350', '預售門票', '預售', 'Advanced', 'Advsnced'
  ]);
  var proofCol = findCol_(header, [
    'Payment Capture / 付款截圖', '付款截圖', 'Payment Capture', '截圖'
  ]);

  var added = 0;
  var skipped = 0;

  for (r = headerRowIdx + 1; r < data.length; r++) {
    var row = data[r];
    var sid = cellStr_(row, sidCol);
    if (!sid || /^submission\s*id$/i.test(sid)) {
      skipped++;
      continue;
    }
    if (done[sid]) {
      skipped++;
      continue;
    }

    var qtyEb = toQty_(cellVal_(row, ebCol));
    var qtyAdv = toQty_(cellVal_(row, advCol));
    var total = toNum_(cellVal_(row, totalCol));
    if (!total) {
      total = qtyEb * TICKET_CFG_.EARLY_BIRD_PRICE + qtyAdv * TICKET_CFG_.ADVANCED_PRICE;
    }

    var summaryParts = [];
    if (qtyEb > 0) summaryParts.push('Early Bird 早鳥x' + qtyEb);
    if (qtyAdv > 0) summaryParts.push('Advanced 預售x' + qtyAdv);
    var productSummary = summaryParts.length ? summaryParts.join(' / ') : '(未選票種)';

    var orderId = 'NA-' + sid;

    // 逐欄賦值，避免 object literal / 貼上損毀導致 "orderProps is not defined"
    var ticketOrder = {};
    ticketOrder.order_id = orderId;
    ticketOrder.submission_id = sid;
    ticketOrder.created_at = cellVal_(row, submittedCol) || new Date();
    ticketOrder.source = TICKET_CFG_.SOURCE;
    ticketOrder.channel = TICKET_CFG_.CHANNEL;
    ticketOrder.customer_name = cellStr_(row, nameCol);
    ticketOrder.email = cellStr_(row, emailCol);
    ticketOrder.phone = cellStr_(row, telCol);
    ticketOrder.metal_pass_id = cellStr_(row, passCol);
    ticketOrder.product_summary = productSummary;
    ticketOrder.qty_early_bird = qtyEb;
    ticketOrder.qty_advanced = qtyAdv;
    ticketOrder.total_payable = total;
    ticketOrder.payment_proof_url = cellStr_(row, proofCol);
    ticketOrder.order_status = '待審核';
    ticketOrder.payment_status = '未付';
    ticketOrder.order_kind = '購票';
    ticketOrder.raw_form_row = 'src|' + sourceName + '|row:' + (r + 1) + '|sid:' + sid;
    ticketOrder.updated_at = new Date();

    appendTicketOrder_(orders, ticketOrder);
    logSheet.appendRow([sid, orderId, new Date(), productSummary, total, ticketOrder.email]);
    done[sid] = true;
    added++;
  }

  SpreadsheetApp.getUi().alert(
    '同步完成\n新增: ' + added + '\n略過: ' + skipped + '\n來源: ' + sourceName
  );
}

/* ========== helpers ========== */

function readSettingsMap_(ss) {
  var sh = ss.getSheetByName(TICKET_CFG_.SETTINGS_SHEET);
  var map = {};
  if (!sh) return map;
  var vals = sh.getDataRange().getValues();
  var i;
  for (i = 1; i < vals.length; i++) {
    var k = String(vals[i][0] || '').trim();
    if (k) map[k] = vals[i][1];
  }
  return map;
}

function ensureSyncLog_(ss) {
  var sh = ss.getSheetByName(TICKET_CFG_.SYNC_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TICKET_CFG_.SYNC_LOG_SHEET);
    sh.appendRow(['submission_id', 'order_id', 'synced_at', 'product_summary', 'total', 'email']);
    sh.hideSheet();
  }
  return sh;
}

function loadSyncedIds_(logSh) {
  var map = {};
  var vals = logSh.getDataRange().getValues();
  var i;
  for (i = 1; i < vals.length; i++) {
    var id = String(vals[i][0] || '').trim();
    if (id) map[id] = true;
  }
  return map;
}

function rowHasAny_(row) {
  var i;
  for (i = 0; i < row.length; i++) {
    if (String(row[i] || '').trim() !== '') return true;
  }
  return false;
}

function nonEmptyJoin_(arr) {
  var out = [];
  var i;
  for (i = 0; i < arr.length; i++) {
    if (arr[i]) out.push(arr[i]);
  }
  return out.join(' | ');
}

function normalizeHeader_(s) {
  var t = String(s || '').toLowerCase();
  t = t.replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, '');
  t = t.replace(/[^a-z0-9\u4e00-\u9fff]+/g, '');
  return t;
}

/** 回傳欄位 index，找不到回 -1 */
function findCol_(header, names) {
  var normHeader = [];
  var i;
  for (i = 0; i < header.length; i++) {
    normHeader.push(normalizeHeader_(header[i]));
  }

  // 精確
  var n;
  for (n = 0; n < names.length; n++) {
    var want = String(names[n] || '').trim();
    for (i = 0; i < header.length; i++) {
      if (header[i] === want) return i;
    }
  }
  // 正規化
  for (n = 0; n < names.length; n++) {
    var nw = normalizeHeader_(names[n]);
    if (!nw) continue;
    for (i = 0; i < normHeader.length; i++) {
      if (normHeader[i] === nw) return i;
    }
  }
  // 包含
  for (n = 0; n < names.length; n++) {
    var needle = normalizeHeader_(names[n]);
    if (!needle || needle.length < 2) continue;
    for (i = 0; i < normHeader.length; i++) {
      if (!normHeader[i]) continue;
      if (normHeader[i].indexOf(needle) >= 0 || needle.indexOf(normHeader[i]) >= 0) {
        return i;
      }
    }
  }
  return -1;
}

function looksLikeSubmissionIdColumn_(data, headerRowIdx, colIdx) {
  var samples = 0;
  var hits = 0;
  var r;
  for (r = headerRowIdx + 1; r < data.length && samples < 8; r++) {
    var v = String(data[r][colIdx] || '').trim();
    if (!v) continue;
    samples++;
    if (/^[A-Za-z0-9_-]{5,16}$/.test(v) && !/^submission/i.test(v)) hits++;
  }
  return samples > 0 && hits >= Math.ceil(samples * 0.6);
}

function cellVal_(row, col) {
  if (col == null || col < 0 || !row) return '';
  return row[col];
}

function cellStr_(row, col) {
  return String(cellVal_(row, col) || '').trim();
}

function toQty_(v) {
  if (v === '' || v == null) return 0;
  var n = Number(v);
  if (!isFinite(n)) return 0;
  return Math.max(0, Math.floor(n));
}

function toNum_(v) {
  if (v === '' || v == null) return 0;
  if (typeof v === 'number') return v;
  var s = String(v).replace(/[^0-9.\-]/g, '');
  var n = Number(s);
  return isFinite(n) ? n : 0;
}

function appendTicketOrder_(ordersSh, ticketOrder) {
  if (!ticketOrder) {
    throw new Error('appendTicketOrder_: ticketOrder 未定義');
  }

  var lastCol = ordersSh.getLastColumn();
  var lastRow = ordersSh.getLastRow();

  // 空表：寫表頭 + 第一筆
  if (lastCol < 1 || lastRow < 1) {
    var defaultHeaders = [
      'order_id', 'created_at', 'source', 'channel', 'customer_name', 'email', 'phone',
      'metal_pass_id', 'product_summary', 'qty_early_bird', 'qty_advanced', 'total_payable',
      'payment_proof_url', 'order_status', 'payment_status', 'order_kind', 'raw_form_row', 'updated_at'
    ];
    ordersSh.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
    ordersSh.appendRow(ticketOrderToArray_(ticketOrder));
    return;
  }

  var header = ordersSh.getRange(1, 1, 1, lastCol).getValues()[0];
  var h = [];
  var i;
  for (i = 0; i < header.length; i++) {
    h.push(String(header[i] || '').trim());
  }

  var aliases = {};
  aliases.order_id = ['order_id', 'Order ID', '訂單編號', 'Submission ID'];
  aliases.submission_id = ['submission_id', 'Submission ID', 'Respondent ID'];
  aliases.created_at = ['created_at', 'Submitted at', '下單時間', 'timestamp'];
  aliases.source = ['source', '來源'];
  aliases.channel = ['channel', '通路'];
  aliases.customer_name = ['customer_name', 'Name/稱呼', 'Name', '姓名', '稱呼', '💀 Name/稱呼'];
  aliases.email = ['email', 'Email/電郵', 'Email', '電郵', '📩 Email/電郵'];
  aliases.phone = ['phone', 'Tel/電話號碼', 'Tel', '電話', '☎️ Tel/電話號碼'];
  aliases.metal_pass_id = ['metal_pass_id', 'Metal Pass 會員編號', 'Metal Pass'];
  aliases.product_summary = ['product_summary', '商品摘要', 'items'];
  aliases.qty_early_bird = ['qty_early_bird', '早鳥', 'Early Bird'];
  aliases.qty_advanced = ['qty_advanced', '預售', 'Advanced'];
  aliases.total_payable = ['total_payable', 'Total Amount', '應付', '總價'];
  aliases.payment_proof_url = ['payment_proof_url', 'Payment Capture / 付款截圖', '付款截圖', 'payment_proof'];
  aliases.order_status = ['order_status', '訂單狀態', 'status'];
  aliases.payment_status = ['payment_status', '付款狀態'];
  aliases.order_kind = ['order_kind', '類型'];
  aliases.raw_form_row = ['raw_form_row', 'raw'];
  aliases.updated_at = ['updated_at', '更新時間'];

  var out = [];
  for (i = 0; i < h.length; i++) {
    out.push('');
  }

  var filled = 0;
  var keys = Object.keys(ticketOrder);
  var k;
  for (k = 0; k < keys.length; k++) {
    var key = keys[k];
    var names = aliases[key] || [key];
    var idx = -1;
    var n;
    for (n = 0; n < names.length; n++) {
      idx = indexOfExact_(h, names[n]);
      if (idx >= 0) break;
    }
    if (idx < 0) {
      // 正規化再找
      for (n = 0; n < names.length; n++) {
        idx = indexOfNorm_(h, names[n]);
        if (idx >= 0) break;
      }
    }
    if (idx >= 0) {
      out[idx] = ticketOrder[key];
      filled++;
    }
  }

  if (filled === 0) {
    // 表頭對不到：附加固定格式
    ordersSh.appendRow(ticketOrderToArray_(ticketOrder));
    return;
  }

  ordersSh.appendRow(out);
}

function ticketOrderToArray_(o) {
  return [
    o.order_id,
    o.created_at,
    o.source,
    o.channel,
    o.customer_name,
    o.email,
    o.phone,
    o.metal_pass_id,
    o.product_summary,
    o.qty_early_bird,
    o.qty_advanced,
    o.total_payable,
    o.payment_proof_url,
    o.order_status,
    o.payment_status,
    o.order_kind,
    o.raw_form_row,
    o.updated_at
  ];
}

function indexOfExact_(arr, name) {
  var i;
  for (i = 0; i < arr.length; i++) {
    if (arr[i] === name) return i;
  }
  return -1;
}

function indexOfNorm_(arr, name) {
  var want = normalizeHeader_(name);
  if (!want) return -1;
  var i;
  for (i = 0; i < arr.length; i++) {
    if (normalizeHeader_(arr[i]) === want) return i;
  }
  return -1;
}
