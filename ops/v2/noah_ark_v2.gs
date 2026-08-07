/**
 * ╔══════════════════════════════════════════════════════════╗
 * ║  Noah's Ark 2.0 — 手機訂單管理 / QR入場 / 分析 / 郵件   ║
 * ╚══════════════════════════════════════════════════════════╝
 *
 * 資料來源（與現有格式對齊）：
 *   NA_Tickets / NA_Merch  ← Tally
 *   Order_Categories       ← tally_column_hint + list_price
 *
 * 產出（AppSheet 主庫）：
 *   v2_Orders    訂單（審核、憑證、寄信、出票）
 *   v2_Tickets   入場票（QR、掃碼狀態）
 *   v2_ScanLog   入場流水
 *   v2_Stats     儀表板數字
 *   v2_Products  可售品項
 *   v2_Chart     Pie 圖（label / sales_qty / revenue）
 *
 * 首次安裝（大試算表易逾時 → 務必分步）：
 *   1) 貼上本檔 → 儲存
 *   2) 執行 installNoahArkV2SheetsOnly   ← 只建空表+表頭（快）
 *   3) 執行 syncProductsFromCategories_
 *   4) 執行 syncTallyToV2Orders
 *   5) 執行 refreshV2StatsAndChart
 *   6) 部署 Web App → Settings.V2_SCAN_WEBAPP_URL
 *
 * 若出現「試算表服務逾時」：
 *   - 關掉其他分頁／AppSheet 同步
 *   - 不要跑 installNoahArkV2（太重），改跑分步
 *   - 每次只跑一個函式
 */

var V2 = {
  NA_TICKETS: 'NA_Tickets',
  NA_MERCH: 'NA_Merch',
  CAT: 'Order_Categories',
  SETTINGS: 'Settings',
  ORDERS: 'v2_Orders',
  TICKETS: 'v2_Tickets',
  SCANLOG: 'v2_ScanLog',
  STATS: 'v2_Stats',
  PRODUCTS: 'v2_Products',
  CHART: 'v2_Chart',
  EVENT_CODE: 'NA2026'
};

/**
 * 讀取 sheet 資料（不含表頭）— 一律用 A1，避免 4 參數 getRange 混淆
 * Apps Script: getRange(row,col,numRows,numCols) ≠ (start,end)
 */
function v2ReadRows_(sh, maxCols) {
  if (!sh) return { headers: [], rows: [] };
  var lastRow = sh.getLastRow();
  var lastCol = Math.min(sh.getLastColumn() || 1, maxCols || 30);
  if (lastRow < 1 || lastCol < 1) return { headers: [], rows: [] };
  var endCol = v2ColA1_(lastCol);
  var headers = sh.getRange('A1:' + endCol + '1').getValues()[0];
  if (lastRow < 2) return { headers: headers, rows: [] };
  var rows = sh.getRange('A2:' + endCol + lastRow).getValues();
  return { headers: headers, rows: rows };
}

function v2ColA1_(col1based) {
  var n = col1based;
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 選單執行時沒有 ss 參數 → 自動取 active spreadsheet */
function v2Ss_(ss) {
  if (ss && typeof ss.getSheetByName === 'function') return ss;
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
}

/* ════════════════════════════════════════
 *  安裝 / 選單
 * ════════════════════════════════════════ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('NOAHSARK-V2')
    .addItem('① 只建 v2 空表（推薦首次）', 'installNoahArkV2SheetsOnly')
    .addItem('①b 安裝+同步（較慢，易逾時）', 'installNoahArkV2')
    .addItem('② 同步 Tally → v2_Orders', 'syncTallyToV2Orders')
    .addItem('②b 同步產品主檔', 'syncProductsFromCategories_')
    .addItem('③ 處理 as_action（審核／出票／寄信）', 'processV2Actions')
    .addItem('④ 刷新統計 + 圖表', 'refreshV2StatsAndChart')
    .addSeparator()
    .addItem('顯示掃碼 Web App 說明', 'showV2ScanDeployHelp')
    .addToUi();
}

/** 只建表頭（無 alert，供其他函式呼叫） */
function createV2SheetsCore_() {
  var ss = v2Ss_();
  ensureSheetHeaders_(ss, V2.ORDERS, v2OrderHeaders_());
  SpreadsheetApp.flush();
  ensureSheetHeaders_(ss, V2.TICKETS, v2TicketHeaders_());
  SpreadsheetApp.flush();
  ensureSheetHeaders_(ss, V2.SCANLOG, v2ScanLogHeaders_());
  SpreadsheetApp.flush();
  ensureSheetHeaders_(ss, V2.STATS, v2StatsHeaders_());
  SpreadsheetApp.flush();
  ensureSheetHeaders_(ss, V2.PRODUCTS, v2ProductHeaders_());
  SpreadsheetApp.flush();
  ensureSheetHeaders_(ss, V2.CHART, v2ChartHeaders_());
  SpreadsheetApp.flush();
  try {
    upsertSetting_(ss, 'V2_ENABLED', 'TRUE');
    upsertSetting_(ss, 'V2_EVENT_CODE', V2.EVENT_CODE);
    upsertSetting_(ss, 'V2_EMAIL_FROM_NAME', "Noah's Ark");
  } catch (e) {}
  return ss;
}

/** 輕量安裝：只建 6 張空表 + 表頭（推薦） */
function installNoahArkV2SheetsOnly() {
  createV2SheetsCore_();
  SpreadsheetApp.getUi().alert(
    'v2 空表已建立（輕量安裝完成）\n\n' +
    '請依序再執行（每次一個，避免逾時）：\n' +
    '1) syncProductsFromCategories_\n' +
    '2) syncTallyToV2Orders\n' +
    '3) refreshV2StatsAndChart'
  );
}

/** 完整安裝（較慢；逾時請改用分步） */
function installNoahArkV2() {
  createV2SheetsCore_();
  SpreadsheetApp.getUi().alert(
    '表頭已建好。\n因檔案很大，請手動依序執行：\n' +
    'syncProductsFromCategories_ → syncTallyToV2Orders → refreshV2StatsAndChart'
  );
}

function ssOrActive_() {
  return SpreadsheetApp.getActive();
}

function v2DailyRun() {
  // 分步 + flush，降低一次鎖死整份檔
  syncTallyToV2Orders();
  SpreadsheetApp.flush();
  processV2Actions();
  SpreadsheetApp.flush();
  refreshV2StatsAndChart();
  SpreadsheetApp.getUi().alert('V2 一鍵完成：同步 + 動作 + 統計');
}

function showV2ScanDeployHelp() {
  SpreadsheetApp.getUi().alert(
    '掃碼 Web App 部署\n\n' +
    '1. Apps Script → 部署 → 新部署\n' +
    '2. 類型：網頁應用程式\n' +
    '3. 執行身分：我\n' +
    '4. 具有存取權的使用者：任何人\n' +
    '5. 複製網址 → Settings 鍵 V2_SCAN_WEBAPP_URL\n\n' +
    '測試：瀏覽器打開\n' +
    'URL?payload=NOAH|NA2026|TEST|TEST|1/1\n\n' +
    'AppSheet 掃碼：AS_Scan 或 v2_ScanLog 寫入 payload 後呼叫 Webhook'
  );
}

/* ════════════════════════════════════════
 *  表頭定義
 * ════════════════════════════════════════ */

function v2OrderHeaders_() {
  return [
    'order_id',           // KEY
    'submission_id',
    'channel',            // ticket | merch
    'channel_zh',         // 票務 | 商品
    'product_label',      // tally_column_hint
    'customer_name',
    'email',
    'phone',
    'metal_pass',
    'sales_qty',
    'unit_price',
    'revenue',
    'payment_proof_url',  // Image
    'payment_status',     // 未付 | 已付
    'order_status',       // 待審核 | 已通過 | 已拒絕 | 已出票
    'receipt_status',     // 未寄出 | 已寄出
    'ticket_ids',         // 出票後填入
    'notes_admin',
    'as_action',          // 通過 | 拒絕 | 寄出郵件 | 出票 | （清空）
    'action_status',      // OK | ERROR | …
    'action_message',
    'created_at',
    'updated_at',
    'reviewed_at',
    'reviewed_by',
    'email_sent_at',
    'source',
    'list_title'          // AppSheet 列表標題
  ];
}

function v2TicketHeaders_() {
  return [
    'ticket_id',          // KEY
    'order_id',
    'qr_payload',         // 掃碼內容
    'qr_image_url',       // Image（QR 圖）
    'customer_name',
    'email',
    'phone',
    'metal_pass',
    'product_label',
    'seq',
    'qty_total',
    'status',             // 有效 | 已入場 | 作廢
    'checked_in_at',
    'checked_in_by',
    'payment_proof_url',
    'as_action',          // 入場
    'action_status',
    'action_message',
    'created_at',
    'updated_at'
  ];
}

function v2ScanLogHeaders_() {
  return [
    'scan_id',            // KEY
    'payload',
    'ticket_id',
    'order_id',
    'customer_name',
    'product_label',
    'result',             // OK | FAIL
    'result_title',
    'message',
    'scanned_by',
    'created_at'
  ];
}

function v2StatsHeaders_() {
  return [
    'metric',             // KEY
    'label_zh',
    'value',
    'unit',
    'sort_order',
    'updated_at'
  ];
}

function v2ProductHeaders_() {
  return [
    'sku',                // KEY
    'channel',
    'product_label',      // tally_column_hint
    'name_zh',
    'list_price',
    'sales_qty',
    'revenue',
    'is_active',
    'sort_order'
  ];
}

function v2ChartHeaders_() {
  return [
    'chart_key',          // KEY
    'channel',
    'label',              // tally_column_hint
    'sales_qty',
    'revenue',
    'updated_at'
  ];
}

/* ════════════════════════════════════════
 *  ② 同步 Tally → v2_Orders
 * ════════════════════════════════════════ */

function syncTallyToV2Orders() {
  var ss = v2Ss_();
  ensureSheetHeaders_(ss, V2.ORDERS, v2OrderHeaders_());
  var catalog = loadCatalogMap_(ss);
  var existing = loadExistingOrderKeys_(ss);

  var ticketLines = readNaTickets_(ss);
  var merchLines = readNaMerch_(ss);

  var ordersSh = ss.getSheetByName(V2.ORDERS);
  var added = 0;
  var now = new Date();
  var i;

  for (i = 0; i < ticketLines.length; i++) {
    var t = ticketLines[i];
    var orderId = 'V2-T-' + t.submission_id;
    if (existing[orderId]) continue;
    var hit = resolveProduct_(t.product_raw, catalog, 'ticket');
    var qty = t.qty;
    var price = hit.price || t.unit_price || 0;
    var rev = qty * price;
    appendOrderRow_(ordersSh, {
      order_id: orderId,
      submission_id: t.submission_id,
      channel: 'ticket',
      channel_zh: '票務',
      product_label: hit.label,
      customer_name: t.customer_name,
      email: t.email,
      phone: t.phone,
      metal_pass: t.metal_pass,
      sales_qty: qty,
      unit_price: price,
      revenue: rev,
      payment_proof_url: t.payment_proof_url,
      payment_status: t.payment_proof_url ? '待核對' : '未付',
      order_status: '待審核',
      receipt_status: '未寄出',
      ticket_ids: '',
      notes_admin: '',
      as_action: '',
      action_status: '',
      action_message: '',
      created_at: t.submitted_at || now,
      updated_at: now,
      reviewed_at: '',
      reviewed_by: '',
      email_sent_at: '',
      source: 'tally_na_tickets',
      list_title: '【待審核】' + (t.customer_name || '') + ' · ' + hit.label + '×' + qty + ' · HK$' + rev
    });
    existing[orderId] = true;
    added++;
  }

  for (i = 0; i < merchLines.length; i++) {
    var m = merchLines[i];
    var oid = 'V2-M-' + m.submission_id + '-' + (m.line_no || 1);
    if (existing[oid]) continue;
    var hitM = resolveProduct_(m.product_raw, catalog, 'merch');
    var qtyM = m.qty;
    var priceM = hitM.price || m.unit_price || 0;
    var revM = qtyM * priceM;
    appendOrderRow_(ordersSh, {
      order_id: oid,
      submission_id: m.submission_id,
      channel: 'merch',
      channel_zh: '商品',
      product_label: hitM.label,
      customer_name: m.customer_name,
      email: m.email,
      phone: m.phone,
      metal_pass: m.metal_pass,
      sales_qty: qtyM,
      unit_price: priceM,
      revenue: revM,
      payment_proof_url: m.payment_proof_url,
      payment_status: m.payment_proof_url ? '待核對' : '未付',
      order_status: '待審核',
      receipt_status: '未寄出',
      ticket_ids: '',
      notes_admin: '',
      as_action: '',
      action_status: '',
      action_message: '',
      created_at: m.submitted_at || now,
      updated_at: now,
      reviewed_at: '',
      reviewed_by: '',
      email_sent_at: '',
      source: 'tally_na_merch',
      list_title: '【待審核】' + (m.customer_name || '') + ' · ' + hitM.label + '×' + qtyM + ' · HK$' + revM
    });
    existing[oid] = true;
    added++;
  }

  // 不在同步後自動刷新統計（易逾時）；請另跑 refreshV2StatsAndChart
  SpreadsheetApp.getUi().alert(
    '同步完成\n新增訂單：' + added +
    '\n票源列：' + ticketLines.length + ' · 商品列：' + merchLines.length +
    '\n\n請再執行：refreshV2StatsAndChart（可選）'
  );
}

function loadExistingOrderKeys_(ss) {
  ss = v2Ss_(ss);
  var sh = ss.getSheetByName(V2.ORDERS);
  var map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  // 只讀 A 欄 order_id（A2:A lastRow）— 用 A1 避免誤讀超大範圍
  var lastRow = sh.getLastRow();
  var vals = sh.getRange('A2:A' + lastRow).getValues();
  var i;
  for (i = 0; i < vals.length; i++) {
    var id = String(vals[i][0] || '').trim();
    if (id) map[id] = true;
  }
  return map;
}

function appendOrderRow_(sh, o) {
  var headers = v2OrderHeaders_();
  var row = [];
  var i;
  for (i = 0; i < headers.length; i++) {
    var k = headers[i];
    row.push(o[k] !== undefined && o[k] !== null ? o[k] : '');
  }
  sh.appendRow(row);
}

/* ════════════════════════════════════════
 *  ③ 處理 as_action
 * ════════════════════════════════════════ */

function processV2Actions() {
  var ss = v2Ss_();
  var sh = ss.getSheetByName(V2.ORDERS);
  if (!sh || sh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('v2_Orders 無資料');
    return;
  }
  ensureSheetHeaders_(ss, V2.TICKETS, v2TicketHeaders_());

  var packed = v2ReadRows_(sh, 30);
  var headers = packed.headers;
  var idx = headerIndexMap_(headers);
  var data = packed.rows;
  var processed = 0;
  var messages = [];
  var r;

  if (idx.as_action == null) {
    SpreadsheetApp.getUi().alert('v2_Orders 缺少 as_action 欄');
    return;
  }

  for (r = 0; r < data.length; r++) {
    var action = String(data[r][idx.as_action] || '').trim();
    if (!action) continue;
    var order = rowToObj_(headers, data[r]);
    var result;
    try {
      if (action === '通過' || action === '確認付款') {
        result = v2ApproveOrder_(ss, order);
      } else if (action === '拒絕') {
        result = v2RejectOrder_(ss, order);
      } else if (action === '出票') {
        result = v2IssueTickets_(ss, order);
      } else if (action === '寄出郵件' || action === '寄信') {
        result = v2SendEmail_(ss, order);
      } else if (action === '已付') {
        result = v2MarkPaid_(ss, order);
      } else {
        result = { ok: false, message: '未知動作: ' + action };
      }
    } catch (err) {
      result = { ok: false, message: String(err.message || err) };
    }

    var sheetRow = r + 2;
    sh.getRange(sheetRow, idx.as_action + 1).setValue('');
    sh.getRange(sheetRow, idx.action_status + 1).setValue(result.ok ? 'OK' : 'ERROR');
    sh.getRange(sheetRow, idx.action_message + 1).setValue(result.message || '');
    sh.getRange(sheetRow, idx.updated_at + 1).setValue(new Date());
    if (result.patch) {
      var p;
      for (p in result.patch) {
        if (idx[p] != null) sh.getRange(sheetRow, idx[p] + 1).setValue(result.patch[p]);
      }
    }
    processed++;
    messages.push(order.order_id + ': ' + (result.message || ''));
  }

  refreshV2StatsAndChart();
  SpreadsheetApp.getUi().alert('處理動作：' + processed + '\n\n' + messages.slice(0, 15).join('\n'));
}

function v2ApproveOrder_(ss, order) {
  ss = v2Ss_(ss);
  var patch = {
    order_status: '已通過',
    payment_status: '已付',
    reviewed_at: new Date(),
    reviewed_by: Session.getActiveUser().getEmail() || 'admin',
    list_title: '【已通過】' + (order.customer_name || '') + ' · ' + order.product_label + ' · HK$' + order.revenue
  };
  // 票務自動出票
  if (String(order.channel) === 'ticket') {
    var issued = v2IssueTickets_(ss, order);
    if (issued.ok) {
      patch.order_status = '已出票';
      patch.ticket_ids = issued.patch && issued.patch.ticket_ids ? issued.patch.ticket_ids : '';
      patch.list_title = '【已出票】' + (order.customer_name || '') + ' · ' + order.product_label;
      return { ok: true, message: '已通過並出票 ' + (issued.message || ''), patch: patch };
    }
    return { ok: true, message: '已通過（出票失敗：' + issued.message + '）', patch: patch };
  }
  return { ok: true, message: '商品訂單已通過', patch: patch };
}

function v2RejectOrder_(ss, order) {
  ss = v2Ss_(ss);
  return {
    ok: true,
    message: '已拒絕',
    patch: {
      order_status: '已拒絕',
      payment_status: '未付',
      reviewed_at: new Date(),
      reviewed_by: Session.getActiveUser().getEmail() || 'admin',
      list_title: '【已拒絕】' + (order.customer_name || '') + ' · ' + order.product_label
    }
  };
}

function v2MarkPaid_(ss, order) {
  ss = v2Ss_(ss);
  return {
    ok: true,
    message: '已標已付',
    patch: { payment_status: '已付', updated_at: new Date() }
  };
}

function v2IssueTickets_(ss, order) {
  ss = v2Ss_(ss);
  ensureSheetHeaders_(ss, V2.TICKETS, v2TicketHeaders_());
  var qty = Math.max(1, toQty_(order.sales_qty));
  var eventCode = getSetting_(ss, 'V2_EVENT_CODE') || V2.EVENT_CODE;
  var tSh = ss.getSheetByName(V2.TICKETS);
  var existing = {};
  if (tSh.getLastRow() >= 2) {
    var lastT = tSh.getLastRow();
    var ids = tSh.getRange(2, 1, lastT, 2).getValues(); // A2:B last — 含 end row
    var i;
    for (i = 0; i < ids.length; i++) {
      if (String(ids[i][1]) === String(order.order_id)) existing[String(ids[i][0])] = true;
    }
  }
  // 已有票則不重複出
  var existingCount = 0;
  for (var k in existing) existingCount++;
  if (existingCount >= qty) {
    return { ok: true, message: '已有 ' + existingCount + ' 張票', patch: { order_status: '已出票' } };
  }

  var start = existingCount + 1;
  var ticketIds = [];
  var now = new Date();
  var seq;
  for (seq = start; seq <= qty; seq++) {
    var ticketId = 'TKT-' + String(order.order_id).replace(/^V2-/, '') + '-' + pad2_(seq);
    var payload = ['NOAH', eventCode, ticketId, order.order_id, seq + '/' + qty].join('|');
    var qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=400x400&margin=12&ecc=H&data=' + encodeURIComponent(payload);
    tSh.appendRow([
      ticketId,
      order.order_id,
      payload,
      qrUrl,
      order.customer_name || '',
      order.email || '',
      order.phone || '',
      order.metal_pass || '',
      order.product_label || '',
      seq,
      qty,
      '有效',
      '',
      '',
      order.payment_proof_url || '',
      '',
      '',
      '',
      now,
      now
    ]);
    ticketIds.push(ticketId);
  }

  // 合併已有 + 新票 id 列表
  var allIds = Object.keys(existing).concat(ticketIds);
  return {
    ok: true,
    message: '出票 ' + ticketIds.length + ' 張',
    patch: {
      order_status: '已出票',
      ticket_ids: allIds.join(', '),
      payment_status: '已付',
      updated_at: now
    }
  };
}

function v2SendEmail_(ss, order) {
  ss = v2Ss_(ss);
  var email = String(order.email || '').trim();
  if (!email || email.indexOf('@') < 0) {
    return { ok: false, message: '無效 email' };
  }
  var name = order.customer_name || '客戶';
  var subject = '【挪亞方舟】訂單通知 ' + order.order_id;
  var body = [];
  body.push(name + ' 您好，');
  body.push('');
  body.push('訂單編號：' + order.order_id);
  body.push('類型：' + (order.channel_zh || order.channel));
  body.push('品項：' + order.product_label);
  body.push('數量：' + order.sales_qty);
  body.push('金額：HK$' + order.revenue);
  body.push('狀態：' + order.order_status + ' / ' + order.payment_status);
  if (order.ticket_ids) body.push('票券：' + order.ticket_ids);
  body.push('');
  body.push('— Noah\'s Ark');
  try {
    GmailApp.sendEmail(email, subject, body.join('\n'), {
      name: getSetting_(ss, 'V2_EMAIL_FROM_NAME') || "Noah's Ark",
      replyTo: getSetting_(ss, 'EMAIL_REPLY_TO') || Session.getActiveUser().getEmail()
    });
    return {
      ok: true,
      message: '已寄出至 ' + email,
      patch: { receipt_status: '已寄出', email_sent_at: new Date() }
    };
  } catch (e) {
    return { ok: false, message: '寄信失敗: ' + e.message };
  }
}

/* ════════════════════════════════════════
 *  ④ 統計 + 圖表
 * ════════════════════════════════════════ */

function refreshV2StatsAndChart() {
  var ss = v2Ss_();
  ensureSheetHeaders_(ss, V2.STATS, v2StatsHeaders_());
  ensureSheetHeaders_(ss, V2.CHART, v2ChartHeaders_());
  syncProductsFromCategories_(ss);

  var orders = readSheetObjects_(ss, V2.ORDERS);
  var tickets = readSheetObjects_(ss, V2.TICKETS);

  var pending = 0, approved = 0, rejected = 0, issued = 0;
  var revTicket = 0, revMerch = 0, qtyTicket = 0, qtyMerch = 0;
  var chartMap = {};
  var i;
  for (i = 0; i < orders.length; i++) {
    var o = orders[i];
    var st = String(o.order_status || '');
    if (st === '待審核') pending++;
    else if (st === '已通過') approved++;
    else if (st === '已拒絕') rejected++;
    else if (st === '已出票') issued++;

    var rev = toMoney_(o.revenue);
    var qty = toQty_(o.sales_qty);
    if (String(o.channel) === 'ticket') {
      revTicket += rev; qtyTicket += qty;
    } else {
      revMerch += rev; qtyMerch += qty;
    }
    var lab = String(o.product_label || '').trim();
    if (lab && st !== '已拒絕') {
      if (!chartMap[lab]) chartMap[lab] = { channel: o.channel, qty: 0, revenue: 0 };
      chartMap[lab].qty += qty;
      chartMap[lab].revenue += rev;
    }
  }

  var validTickets = 0, checkedIn = 0;
  for (i = 0; i < tickets.length; i++) {
    var t = tickets[i];
    if (String(t.status) === '有效') validTickets++;
    if (String(t.status) === '已入場') checkedIn++;
  }

  var now = new Date();
  var stats = [
    ['pending_orders', '待審核訂單', pending, '單', 1, now],
    ['approved_orders', '已通過訂單', approved, '單', 2, now],
    ['issued_orders', '已出票訂單', issued, '單', 3, now],
    ['rejected_orders', '已拒絕訂單', rejected, '單', 4, now],
    ['revenue_ticket', '票務銷售額', revTicket, 'HKD', 10, now],
    ['revenue_merch', '商品銷售額', revMerch, 'HKD', 11, now],
    ['revenue_total', '總銷售額', revTicket + revMerch, 'HKD', 12, now],
    ['qty_ticket', '票務銷量', qtyTicket, '張', 20, now],
    ['qty_merch', '商品銷量', qtyMerch, '件', 21, now],
    ['tickets_valid', '有效入場票', validTickets, '張', 30, now],
    ['tickets_checked_in', '已入場', checkedIn, '人', 31, now]
  ];
  writeSheetReplace_(ss, V2.STATS, v2StatsHeaders_(), stats, '#d0e2ff');

  var chartRows = [];
  var labs = Object.keys(chartMap);
  labs.sort(function (a, b) { return chartMap[b].revenue - chartMap[a].revenue; });
  for (i = 0; i < labs.length; i++) {
    var c = chartMap[labs[i]];
    chartRows.push([
      'CH-' + (i + 1),
      c.channel || '',
      labs[i],
      c.qty,
      c.revenue,
      now
    ]);
  }
  writeSheetReplace_(ss, V2.CHART, v2ChartHeaders_(), chartRows, '#fff2cc');
}

function syncProductsFromCategories_(ss) {
  ss = v2Ss_(ss);
  ensureSheetHeaders_(ss, V2.PRODUCTS, v2ProductHeaders_());
  var cat = ss.getSheetByName(V2.CAT);
  if (!cat || cat.getLastRow() < 2) return;
  var packedCat = v2ReadRows_(cat, 25);
  var headers = packedCat.headers;
  var idx = headerIndexMap_(headers);
  var data = packedCat.rows;
  var rows = [];
  var r;
  for (r = 0; r < data.length; r++) {
    if (idx.level != null && Number(data[r][idx.level]) !== 3) continue;
    var hint = cell_(data[r], idx, 'tally_column_hint') || cell_(data[r], idx, 'form_option_label') || cell_(data[r], idx, 'name_zh');
    var sku = cell_(data[r], idx, 'sku') || ('SKU-' + (r + 1));
    rows.push([
      sku,
      cell_(data[r], idx, 'channel'),
      hint,
      cell_(data[r], idx, 'name_zh'),
      toMoney_(idx.list_price != null ? data[r][idx.list_price] : 0),
      toQty_(idx['銷售數量'] != null ? data[r][idx['銷售數量']] : 0),
      toMoney_(idx['單件貨品總收入'] != null ? data[r][idx['單件貨品總收入']] : 0),
      cell_(data[r], idx, 'is_active') || 'TRUE',
      data[r][idx.sort_order] || 0
    ]);
  }
  writeSheetReplace_(ss, V2.PRODUCTS, v2ProductHeaders_(), rows, '#e6f4ea');
}

/* ════════════════════════════════════════
 *  Web App：掃碼入場
 * ════════════════════════════════════════ */

function doGet(e) {
  return handleScan_(e);
}

function doPost(e) {
  return handleScan_(e);
}

function handleScan_(e) {
  e = e || {};
  var p = e.parameter || {};
  var payload = p.payload || p.qr || p.data || '';
  if (!payload && e.postData && e.postData.contents) {
    try {
      var j = JSON.parse(e.postData.contents);
      payload = j.payload || j.qr || j.data || '';
    } catch (err) {
      payload = e.postData.contents;
    }
  }
  payload = String(payload || '').trim();
  var scannedBy = p.by || p.user || 'gate';

  var result = processScanPayload_(payload, scannedBy);
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function processScanPayload_(payload, scannedBy) {
  var ss = v2Ss_();
  ensureSheetHeaders_(ss, V2.SCANLOG, v2ScanLogHeaders_());
  ensureSheetHeaders_(ss, V2.TICKETS, v2TicketHeaders_());

  var scanId = 'SCN-' + Utilities.getUuid().slice(0, 8);
  var now = new Date();
  var base = {
    scan_id: scanId,
    payload: payload,
    ticket_id: '',
    order_id: '',
    customer_name: '',
    product_label: '',
    result: 'FAIL',
    result_title: '失敗',
    message: '',
    scanned_by: scannedBy || 'gate',
    created_at: now
  };

  if (!payload) {
    base.message = '空 payload';
    appendScanLog_(ss, base);
    return base;
  }

  // 支援完整 QR 或純 ticket_id
  var ticketId = payload;
  if (payload.indexOf('|') >= 0) {
    var parts = payload.split('|');
    // NOAH|EVENT|ticketId|orderId|seq/total
    if (parts.length >= 3) ticketId = parts[2];
  }

  var tSh = ss.getSheetByName(V2.TICKETS);
  if (!tSh || tSh.getLastRow() < 2) {
    base.message = '無票庫';
    appendScanLog_(ss, base);
    return base;
  }

  var packedT = v2ReadRows_(tSh, 20);
  var headers = packedT.headers;
  var idx = headerIndexMap_(headers);
  var data = packedT.rows;
  var r;
  var found = -1;
  for (r = 0; r < data.length; r++) {
    var id = String(data[r][idx.ticket_id] || '');
    var pay = String(data[r][idx.qr_payload] || '');
    if (id === ticketId || pay === payload || id === payload) {
      found = r;
      break;
    }
  }

  if (found < 0) {
    base.message = '找不到票券';
    base.result_title = '無效 QR';
    appendScanLog_(ss, base);
    return base;
  }

  var row = data[found];
  var status = String(row[idx.status] || '');
  base.ticket_id = String(row[idx.ticket_id] || '');
  base.order_id = String(row[idx.order_id] || '');
  base.customer_name = String(row[idx.customer_name] || '');
  base.product_label = String(row[idx.product_label] || '');

  if (status === '作廢') {
    base.message = '票已作廢';
    base.result_title = '已作廢';
    appendScanLog_(ss, base);
    return base;
  }
  if (status === '已入場') {
    base.message = '重複掃碼 · 原入場 ' + String(row[idx.checked_in_at] || '');
    base.result_title = '已入場';
    base.result = 'FAIL';
    appendScanLog_(ss, base);
    return base;
  }

  // 標記入場
  var sheetRow = found + 2;
  tSh.getRange(sheetRow, idx.status + 1).setValue('已入場');
  tSh.getRange(sheetRow, idx.checked_in_at + 1).setValue(now);
  tSh.getRange(sheetRow, idx.checked_in_by + 1).setValue(scannedBy);
  tSh.getRange(sheetRow, idx.updated_at + 1).setValue(now);

  base.result = 'OK';
  base.result_title = 'CONFIRMED';
  base.message = base.customer_name + ' · ' + base.product_label + ' · 入場成功';
  appendScanLog_(ss, base);
  return base;
}

function appendScanLog_(ss, o) {
  ss = v2Ss_(ss);
  var sh = ss.getSheetByName(V2.SCANLOG);
  var headers = v2ScanLogHeaders_();
  var row = [];
  var i;
  for (i = 0; i < headers.length; i++) {
    row.push(o[headers[i]] !== undefined ? o[headers[i]] : '');
  }
  sh.appendRow(row);
}

/**
 * AppSheet 也可呼叫：在 v2_Tickets 設 as_action=入場
 */
function processV2TicketGateActions() {
  var ss = v2Ss_();
  var sh = ss.getSheetByName(V2.TICKETS);
  if (!sh || sh.getLastRow() < 2) return;
  var packedG = v2ReadRows_(sh, 20);
  var headers = packedG.headers;
  var idx = headerIndexMap_(headers);
  if (idx.as_action == null) return;
  var data = packedG.rows;
  var r;
  for (r = 0; r < data.length; r++) {
    var action = String(data[r][idx.as_action] || '').trim();
    if (action !== '入場') continue;
    var payload = String(data[r][idx.qr_payload] || data[r][idx.ticket_id] || '');
    var res = processScanPayload_(payload, Session.getActiveUser().getEmail() || 'appsheet');
    sh.getRange(r + 2, idx.as_action + 1).setValue('');
    if (idx.action_status != null) sh.getRange(r + 2, idx.action_status + 1).setValue(res.result);
    if (idx.action_message != null) sh.getRange(r + 2, idx.action_message + 1).setValue(res.message);
  }
}

/* ════════════════════════════════════════
 *  讀 NA_Tickets / NA_Merch
 * ════════════════════════════════════════ */

function readNaTickets_(ss) {
  ss = v2Ss_(ss);
  var sh = ss.getSheetByName(V2.NA_TICKETS);
  var out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var packed = v2ReadRows_(sh, 16);
  var headers = packed.headers;
  var data = packed.rows;
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
  var r, c;

  if (cType != null && cQty != null) {
    for (r = 0; r < data.length; r++) {
      var type = String(data[r][cType] || '').trim();
      var qty = toQty_(data[r][cQty]);
      if (!type || qty <= 0) continue;
      var sid = cSid != null ? String(data[r][cSid] || '').trim() : ('R' + (r + 2));
      out.push({
        submission_id: sid || ('R' + (r + 2)),
        submitted_at: cAt != null ? data[r][cAt] : '',
        customer_name: cName != null ? data[r][cName] : '',
        email: cEmail != null ? data[r][cEmail] : '',
        phone: cTel != null ? data[r][cTel] : '',
        metal_pass: cPass != null ? data[r][cPass] : '',
        product_raw: type,
        qty: qty,
        unit_price: cPrice != null ? safePrice_(data[r][cPrice]) : 0,
        payment_proof_url: normalizeProofUrl_(cProof != null ? data[r][cProof] : '')
      });
    }
    return out;
  }
  for (c = 0; c < headers.length; c++) {
    var hh = String(headers[c] || '').trim();
    if (!/早鳥|預售|會員|Metal|Early|Advanced|門票/i.test(hh)) continue;
    for (r = 0; r < data.length; r++) {
      var q = toQty_(data[r][c]);
      if (q <= 0) continue;
      var sid2 = cSid != null ? String(data[r][cSid] || '').trim() : ('R' + (r + 2));
      out.push({
        submission_id: sid2,
        submitted_at: cAt != null ? data[r][cAt] : '',
        customer_name: cName != null ? data[r][cName] : '',
        email: cEmail != null ? data[r][cEmail] : '',
        phone: cTel != null ? data[r][cTel] : '',
        metal_pass: cPass != null ? data[r][cPass] : '',
        product_raw: hh,
        qty: q,
        unit_price: 0,
        payment_proof_url: normalizeProofUrl_(cProof != null ? data[r][cProof] : '')
      });
    }
  }
  return out;
}

function readNaMerch_(ss) {
  ss = v2Ss_(ss);
  var sh = ss.getSheetByName(V2.NA_MERCH);
  var out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var packed = v2ReadRows_(sh, 20);
  var headers = packed.headers;
  var data = packed.rows;
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
  var r, c;
  var line = 0;
  var seen = {};

  if (cCat != null && cQty != null) {
    for (r = 0; r < data.length; r++) {
      var cat = String(data[r][cCat] || '').trim();
      var qty = toQty_(data[r][cQty]);
      if (!cat || qty <= 0) continue;
      var sid = cSid != null ? String(data[r][cSid] || '').trim() : ('R' + (r + 2));
      line++;
      var uk = sid + '|' + cat + '|' + qty + '|ijk';
      if (seen[uk]) continue;
      seen[uk] = true;
      out.push({
        submission_id: sid,
        line_no: line,
        submitted_at: cAt != null ? data[r][cAt] : '',
        customer_name: cName != null ? data[r][cName] : '',
        email: cEmail != null ? data[r][cEmail] : '',
        phone: cTel != null ? data[r][cTel] : '',
        metal_pass: cPass != null ? data[r][cPass] : '',
        product_raw: cat,
        qty: qty,
        unit_price: cPrice != null ? safePrice_(data[r][cPrice]) : 0,
        payment_proof_url: normalizeProofUrl_(cProof != null ? data[r][cProof] : '')
      });
    }
  }

  for (c = 0; c < headers.length; c++) {
    var hh = String(headers[c] || '').trim();
    if (!isMerchProductHeader_(hh)) continue;
    if (c === cCat || c === cQty || c === cPrice) continue;
    for (r = 0; r < data.length; r++) {
      var q = toQty_(data[r][c]);
      if (q <= 0) continue;
      var sid2 = cSid != null ? String(data[r][cSid] || '').trim() : ('R' + (r + 2));
      var uk2 = sid2 + '|' + hh + '|' + q + '|col';
      var ukI = sid2 + '|' + hh + '|' + q + '|ijk';
      if (seen[uk2] || seen[ukI]) continue;
      seen[uk2] = true;
      line++;
      out.push({
        submission_id: sid2,
        line_no: line,
        submitted_at: cAt != null ? data[r][cAt] : '',
        customer_name: cName != null ? data[r][cName] : '',
        email: cEmail != null ? data[r][cEmail] : '',
        phone: cTel != null ? data[r][cTel] : '',
        metal_pass: cPass != null ? data[r][cPass] : '',
        product_raw: hh,
        qty: q,
        unit_price: 0,
        payment_proof_url: normalizeProofUrl_(cProof != null ? data[r][cProof] : '')
      });
    }
  }
  return out;
}

/* ════════════════════════════════════════
 *  Catalog / utils
 * ════════════════════════════════════════ */

function loadCatalogMap_(ss) {
  ss = v2Ss_(ss);
  var map = {};
  var sh = ss.getSheetByName(V2.CAT);
  if (!sh || sh.getLastRow() < 2) return map;
  var packed = v2ReadRows_(sh, 25);
  var headers = packed.headers;
  var idx = headerIndexMap_(headers);
  var data = packed.rows;
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
    regCat_(map, label, price, channel);
    regCat_(map, hint, price, channel);
    regCat_(map, form, price, channel);
    regCat_(map, nameZh, price, channel);
    regCat_(map, stripEmoji_(label), price, channel);
  }
  return map;
}

function regCat_(map, label, price, channel) {
  if (!label) return;
  var L = String(label).trim();
  var k = normKey_(L);
  if (!k) return;
  if (!map[k] || L.length >= map[k].label.length) {
    map[k] = { label: L, price: price, channel: channel || '' };
  }
}

function resolveProduct_(raw, map, preferChannel) {
  var s = String(raw || '').trim();
  var fb = { label: s || '(未分類)', price: guessPrice_(s), channel: preferChannel || '' };
  if (!s) return fb;
  var k = normKey_(s);
  var k2 = normKey_(stripEmoji_(s));
  if (map[k]) return map[k];
  if (map[k2]) return map[k2];
  var keys = Object.keys(map);
  var i;
  for (i = 0; i < keys.length; i++) {
    var item = map[keys[i]];
    if (preferChannel && item.channel && item.channel !== preferChannel) continue;
    if (k.indexOf(keys[i]) >= 0 || keys[i].indexOf(k) >= 0) return item;
  }
  if (/早鳥|early/i.test(s)) {
    for (i = 0; i < keys.length; i++) if (/早鳥|early/i.test(map[keys[i]].label)) return map[keys[i]];
  }
  if (/預售|advanced|presale/i.test(s)) {
    for (i = 0; i < keys.length; i++) if (/預售|advanced|presale/i.test(map[keys[i]].label)) return map[keys[i]];
  }
  if (/會員|metal|特級/i.test(s)) {
    for (i = 0; i < keys.length; i++) if (/會員|metal|特級/i.test(map[keys[i]].label)) return map[keys[i]];
  }
  return fb;
}

function ensureSheetHeaders_(ss, name, headers) {
  ss = v2Ss_(ss);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  // 只在空白表寫表頭；已有資料不 insertRows（防逾時）
  var a1 = '';
  try {
    a1 = String(sh.getRange('A1').getValue() || '');
  } catch (e) {
    a1 = '';
  }
  if (!a1) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    try {
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#cfe2f3');
      sh.setFrozenRows(1);
    } catch (e2) {}
  }
  return sh;
}

function writeSheetReplace_(ss, name, headers, body, color) {
  ss = v2Ss_(ss);
  var sh = ensureSheetHeaders_(ss, name, headers);
  var cols = headers.length;
  var nBody = body ? body.length : 0;
  var total = 1 + nBody;

  // 小表寫入：只清「新資料需要的列數 + 少量舊列」，最多 120 列（防逾時）
  var oldLast = sh.getLastRow();
  var clearTo = Math.min(Math.max(oldLast, total), total + 20, 120);
  var clearCols = Math.min(Math.max(sh.getLastColumn(), cols), cols + 2, 20);
  try {
    if (clearTo >= 1 && clearCols >= 1) {
      sh.getRange(1, 1, clearTo, clearCols).clearContent();
    }
  } catch (e) {}

  var all = [headers];
  var i;
  for (i = 0; i < nBody; i++) all.push(body[i]);
  // setValues：起點 (1,1)，列數 all.length，欄數 cols
  sh.getRange(1, 1, all.length, cols).setValues(all);
  try {
    sh.getRange(1, 1, 1, cols).setFontWeight('bold').setBackground(color || '#cfe2f3');
    sh.setFrozenRows(1);
  } catch (e2) {}

  if (nBody > 0) {
    var hi = headerIndexMap_(headers);
    ['sales_qty', 'value', 'unit_price', 'revenue', 'seq', 'qty_total', 'sort_order'].forEach(function (colName) {
      if (hi[colName] != null) {
        try {
          // 數字格式：從第2列、該欄，nBody 列 × 1 欄
          sh.getRange(2, hi[colName] + 1, nBody, 1).setNumberFormat('0');
        } catch (e3) {}
      }
    });
  }
}

function readSheetObjects_(ss, name) {
  ss = v2Ss_(ss);
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var packed = v2ReadRows_(sh, 30);
  var headers = packed.headers;
  var data = packed.rows;
  var out = [];
  var r;
  for (r = 0; r < data.length; r++) {
    if (!String(data[r][0] || '').trim()) continue;
    out.push(rowToObj_(headers, data[r]));
  }
  return out;
}

function rowToObj_(headers, row) {
  var o = {};
  var i;
  for (i = 0; i < headers.length; i++) {
    o[String(headers[i] || '').trim()] = row[i];
  }
  return o;
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

function cell_(row, idx, name) {
  if (!idx || idx[name] == null) return '';
  return String(row[idx[name]] || '').trim();
}

function upsertSetting_(ss, key, value) {
  ss = v2Ss_(ss);
  var sh = ss.getSheetByName(V2.SETTINGS);
  if (!sh) return;
  // 只讀 A:B，避免 getDataRange 掃到整份超大區域
  var lastRow = Math.min(sh.getLastRow() || 1, 80);
  var vals = sh.getRange(1, 1, lastRow, 2).getValues();
  var i;
  for (i = 1; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

function getSetting_(ss, key) {
  ss = v2Ss_(ss);
  var sh = ss.getSheetByName(V2.SETTINGS);
  if (!sh) return '';
  var lastRow = Math.min(sh.getLastRow() || 1, 80);
  var vals = sh.getRange(1, 1, lastRow, 2).getValues();
  var i;
  for (i = 1; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() === key) return vals[i][1];
  }
  return '';
}

function normalizeProofUrl_(v) {
  if (v === '' || v == null) return '';
  var s = String(v).trim();
  var m = s.match(/https?:\/\/[^\s,;"']+/);
  return m ? m[0] : s;
}

function isMerchProductHeader_(h) {
  if (/Submission|Respondent|Submitted|Total|Name|Email|Tel|Metal Pass|付款|Payment|數量|單價|商品分欄|ID/i.test(h)) return false;
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
  if (typeof v === 'number') return isFinite(v) && v < 100000 ? v : 0;
  var n = Number(String(v).replace(/[^0-9.\-]/g, ''));
  return isFinite(n) && n < 100000 ? n : 0;
}

function safePrice_(v) {
  var n = toMoney_(v);
  return n > 0 && n < 10000 ? n : 0;
}

function pad2_(n) {
  return n < 10 ? '0' + n : String(n);
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
