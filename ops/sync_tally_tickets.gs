/**
 * Noah's Ark — 同步 Tally 購票表單 → Orders
 *
 * 來源分頁表頭（Tally 自動寫入）:
 * Submission ID | Respondent ID | Submitted at | Total Amount |
 * 💀 Name/稱呼 | 📩 Email/電郵 | ☎️ Tel/電話號碼 | Metal Pass 會員編號 |
 * Untitled number field | 🎫 早鳥門票 HKD300 | 🎫 預售 HKD350 |
 * Payment Capture / 付款截圖
 *
 * 使用方式:
 * 1) 貼到試算表 Apps Script 專案（可與現有腳本同專案）
 * 2) Settings!TICKET_SOURCE_TAB_NAME = 你的分頁名（例如 NA_Tickets）
 * 3) 選單 NOAHSARK → 同步 Tally 購票 → Orders
 *
 * 注意: 若你專案已有 onOpen / 選單，請只合併 menu 項目，避免重複 onOpen。
 */

var TICKET_CFG_ = {
  SETTINGS_SHEET: 'Settings',
  ORDERS_SHEET: 'Orders',
  // 來源分頁優先讀 Settings.TICKET_SOURCE_TAB_NAME，否則用此預設
  DEFAULT_SOURCE_TAB: 'NA_Tickets',
  // 已同步過的 Submission ID 記錄分頁（自動建立）
  SYNC_LOG_SHEET: '_TallyTicketSyncLog',
  EARLY_BIRD_PRICE: 300,
  ADVANCED_PRICE: 350,
  SOURCE: 'ticket_form',
  CHANNEL: 'tally',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('NOAHSARK')
    .addItem('同步 Tally 購票 → Orders', 'syncTallyTicketsToOrders')
    .addToUi();
}

/** 手動／時間觸發入口 */
function syncTallyTicketsToOrders() {
  var ss = SpreadsheetApp.getActive();
  var settings = readSettingsMap_(ss);
  var sourceName = settings.TICKET_SOURCE_TAB_NAME || TICKET_CFG_.DEFAULT_SOURCE_TAB;
  var source = ss.getSheetByName(sourceName);
  if (!source) {
    throw new Error('找不到來源分頁: ' + sourceName + '（請在 Settings 設 TICKET_SOURCE_TAB_NAME）');
  }

  var orders = ss.getSheetByName(TICKET_CFG_.ORDERS_SHEET);
  if (!orders) throw new Error('找不到 Orders 分頁');

  var log = ensureSyncLog_(ss);
  var done = loadSyncedIds_(log);

  var data = source.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('來源分頁沒有資料列');
    return;
  }

  // 跳過完全空白的表頭列（有時 Tally 第一列不是表頭）
  var headerRowIdx = 0;
  while (headerRowIdx < Math.min(5, data.length)) {
    var probe = data[headerRowIdx].some(function (c) { return String(c || '').trim() !== ''; });
    if (probe) break;
    headerRowIdx++;
  }
  var header = data[headerRowIdx].map(function (h) { return String(h || '').trim(); });
  var col = indexMapFlexible_(header);

  // 必要欄（模糊匹配：Submission ID / submission_id / 提交 ID 等）
  var sidKey = firstExistingFlexible_(col, [
    'Submission ID', 'submission id', 'submission_id', 'SubmissionID',
    '提交 ID', '提交ID', '提交編號', 'Response ID', 'response_id',
  ]);
  var submittedKey = firstExistingFlexible_(col, [
    'Submitted at', 'submitted at', 'submitted_at', 'Submitted At',
    '提交時間', '提交於', 'Created at', 'created_at', 'Timestamp',
  ]);
  var totalKey = firstExistingFlexible_(col, [
    'Total Amount', 'total amount', 'total_amount', 'Total',
    '總金額', '總價', '應付', 'Amount',
  ]);

  if (sidKey == null) {
    // 後備：第 1 欄若看起來像 Tally ID（短英數），就當 Submission ID
    if (header.length > 0 && looksLikeSubmissionIdColumn_(data, headerRowIdx, 0)) {
      sidKey = 0;
    } else {
      throw new Error(
        '來源分頁找不到 Submission ID 欄。\n' +
        '實際表頭（第 ' + (headerRowIdx + 1) + ' 列）:\n' +
        header.filter(Boolean).join(' | ') +
        '\n\n請確認 Settings.TICKET_SOURCE_TAB_NAME 指向 Tally 寫入的分頁，' +
        '且第一列包含 Submission ID（或把第一欄改名為 Submission ID）。'
      );
    }
  }
  if (submittedKey == null) {
    submittedKey = firstExistingFlexible_(col, ['Date', '日期', '時間']) ;
  }
  if (totalKey == null) {
    // 允許稍後用票價推算
    totalKey = null;
  }

  var nameKey = firstExistingFlexible_(col, [
    '💀 Name/稱呼', 'Name/稱呼', 'Name', '稱呼', '姓名', 'name',
  ]);
  var emailKey = firstExistingFlexible_(col, [
    '📩 Email/電郵', 'Email/電郵', 'Email', '電郵', 'email', 'E-mail',
  ]);
  var telKey = firstExistingFlexible_(col, [
    '☎️ Tel/電話號碼', 'Tel/電話號碼', 'Tel', '電話', '電話號碼', 'Phone', 'phone',
  ]);
  var passKey = firstExistingFlexible_(col, [
    'Metal Pass 會員編號', 'Metal Pass', 'METAL-PASS', 'metal pass', '會員編號',
  ]);
  var ebKey = firstExistingFlexible_(col, [
    '🎫 早鳥門票 HKD300', '早鳥門票 HKD300', '早鳥門票', '早鳥', 'Early Bird', 'EarlyBird',
  ]);
  var advKey = firstExistingFlexible_(col, [
    '🎫 預售 HKD350', '預售 HKD350', '預售門票', '預售', 'Advanced', 'Advsnced',
  ]);
  var proofKey = firstExistingFlexible_(col, [
    'Payment Capture / 付款截圖', '付款截圖', 'Payment Capture', 'payment capture', '截圖',
  ]);

  var added = 0;
  var skipped = 0;

  for (var r = headerRowIdx + 1; r < data.length; r++) {
    var row = data[r];
    var sid = String(row[sidKey] || '').trim();
    // 跳過表頭重列或合計列
    if (!sid || /^submission\s*id$/i.test(sid)) { skipped++; continue; }
    if (done[sid]) { skipped++; continue; }

    var name = nameKey != null ? String(row[nameKey] || '').trim() : '';
    var email = emailKey != null ? String(row[emailKey] || '').trim() : '';
    var phone = telKey != null ? String(row[telKey] || '').trim() : '';
    var metal = passKey != null ? String(row[passKey] || '').trim() : '';
    var qtyEb = toQty_(ebKey != null ? row[ebKey] : 0);
    var qtyAdv = toQty_(advKey != null ? row[advKey] : 0);
    var total = totalKey != null ? toNum_(row[totalKey]) : 0;
    var proof = proofKey != null ? String(row[proofKey] || '').trim() : '';
    var submitted = submittedKey != null ? row[submittedKey] : '';

    // 若 Total Amount 空白，用單價推算
    if (!total) {
      total = qtyEb * TICKET_CFG_.EARLY_BIRD_PRICE + qtyAdv * TICKET_CFG_.ADVANCED_PRICE;
    }

    var summaryParts = [];
    if (qtyEb > 0) summaryParts.push('Early Bird 早鳥×' + qtyEb);
    if (qtyAdv > 0) summaryParts.push('Advanced 預售×' + qtyAdv);
    var productSummary = summaryParts.join('、') || '（未選票種）';

    var orderId = 'NA-' + sid; // 穩定、可對回 Tally Submission ID
    // 若你現有腳本用 ORD- 格式，可改呼叫既有 generateOrderId_()

    // ---- 寫入 Orders：採「附加列」通用欄位 ----
    // 若你的 Orders 表頭與下方不同，請改 map 到你真實欄位（見腳本底部 note）
    var orderRow = {
      order_id: orderId,
      submission_id: sid,
      created_at: submitted || new Date(),
      source: TICKET_CFG_.SOURCE,
      channel: TICKET_CFG_.CHANNEL,
      customer_name: name,
      email: email,
      phone: phone,
      metal_pass_id: metal,
      product_summary: productSummary,
      qty_early_bird: qtyEb,
      qty_advanced: qtyAdv,
      total_payable: total,
      payment_proof_url: proof,
      order_status: '待審核',
      payment_status: '未付',
      order_kind: '購票',
      raw_form_row: 'src|' + sourceName + '|row:' + (r + 1) + '|sid:' + sid,
      updated_at: new Date(),
    };

    appendOrderByHeader_(orders, orderProps);
    log.appendRow([sid, orderId, new Date(), productSummary, total, email]);
    done[sid] = true;
    added++;
  }

  SpreadsheetApp.getUi().alert(
    '同步完成\n新增: ' + added + '\n略過（已同步／空）: ' + skipped + '\n來源: ' + sourceName
  );
}

/* ================= helpers ================= */

function readSettingsMap_(ss) {
  var sh = ss.getSheetByName(TICKET_CFG_.SETTINGS_SHEET);
  var map = {};
  if (!sh) return map;
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    var k = String(vals[i][0] || '').trim();
    var v = vals[i][1];
    if (k) map[k] = v;
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
  for (var i = 1; i < vals.length; i++) {
    var id = String(vals[i][0] || '').trim();
    if (id) map[id] = true;
  }
  return map;
}

function indexMap_(header) {
  var m = {};
  header.forEach(function (h, i) { m[h] = i; });
  return m;
}

/** 正規化表頭：小寫、去空白／符號，方便模糊比對（Apps Script V8 相容） */
function normalizeHeader_(s) {
  var t = String(s || '').toLowerCase();
  // 去掉常見 emoji 區段（用代理對，避免 \u{} 相容問題）
  t = t.replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, '');
  // 只留 a-z 0-9 與中文
  t = t.replace(/[^a-z0-9\u4e00-\u9fff]+/g, '');
  return t;
}

/**
 * col 同時保留：
 *  - 原始表頭 → index
 *  - 正規化表頭 → index
 */
function indexMapFlexible_(header) {
  var m = {};
  header.forEach(function (h, i) {
    var raw = String(h || '').trim();
    if (!raw) return;
    m[raw] = i;
    m[normalizeHeader_(raw)] = i;
  });
  return m;
}

function firstExisting_(col, names) {
  for (var i = 0; i < names.length; i++) {
    if (col[names[i]] != null) return col[names[i]];
  }
  return null;
}

function firstExistingFlexible_(col, names) {
  // 1) 精確
  var hit = firstExisting_(col, names);
  if (hit != null) return hit;
  // 2) 正規化後比對
  for (var i = 0; i < names.length; i++) {
    var n = normalizeHeader_(names[i]);
    if (n && col[n] != null) return col[n];
  }
  // 3) 包含關係（表頭含關鍵字）
  var keys = Object.keys(col);
  for (var j = 0; j < names.length; j++) {
    var needle = normalizeHeader_(names[j]);
    if (!needle || needle.length < 2) continue;
    for (var k = 0; k < keys.length; k++) {
      var key = keys[k];
      // 只對正規化 key 做 includes，避免重複
      if (key !== normalizeHeader_(key)) continue;
      if (key.indexOf(needle) >= 0 || needle.indexOf(key) >= 0) {
        return col[key];
      }
    }
  }
  return null;
}

/** 抽樣判斷某欄是否像 Tally Submission ID（短英數混合） */
function looksLikeSubmissionIdColumn_(data, headerRowIdx, colIdx) {
  var samples = 0;
  var hits = 0;
  for (var r = headerRowIdx + 1; r < data.length && samples < 8; r++) {
    var v = String(data[r][colIdx] || '').trim();
    if (!v) continue;
    samples++;
    // Tally ID 常見：6–12 位英數，如 WJDZ4Kj / PR80LgP
    if (/^[A-Za-z0-9_-]{5,16}$/.test(v) && !/^submission/i.test(v)) hits++;
  }
  return samples > 0 && hits >= Math.ceil(samples * 0.6);
}

function toQty_(v) {
  if (v === '' || v == null) return 0;
  var n = Number(v);
  return isFinite(n) ? Math.max(0, Math.floor(n)) : 0;
}

function toNum_(v) {
  if (v === '' || v == null) return 0;
  if (typeof v === 'number') return v;
  var s = String(v).replace(/[^0-9.\-]/g, '');
  var n = Number(s);
  return isFinite(n) ? n : 0;
}

/**
 * 依 Orders 第一列表頭寫入一列。
 * 支援兩種常見 schema：
 *  A) 正規化欄位 order_id, source, customer_name...
 *  B) 舊式 / 混用欄位（盡量智能對應）
 */
function appendOrderByHeader_(ordersSh, props) {
  var header = ordersSh.getRange(1, 1, 1, ordersSh.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h || '').trim(); });

  // 若 Orders 是空表或只有標題像 merch 原始匯入，改寫「審核友善」最小集
  var aliases = {
    order_id: ['order_id', 'Order ID', '訂單編號', 'Submission ID'],
    submission_id: ['submission_id', 'Submission ID', 'Respondent ID'],
    created_at: ['created_at', 'Submitted at', '下單時間', 'timestamp'],
    source: ['source', '來源'],
    channel: ['channel', '通路'],
    customer_name: ['customer_name', '💀 Name/稱呼', 'Name', '姓名', '稱呼'],
    email: ['email', '📩 Email/電郵', 'Email', '電郵'],
    phone: ['phone', '☎️ Tel/電話號碼', 'Tel', '電話'],
    metal_pass_id: ['metal_pass_id', 'Metal Pass 會員編號', 'Metal Pass'],
    product_summary: ['product_summary', '商品摘要', 'items'],
    total_payable: ['total_payable', 'Total Amount', '應付', '總價'],
    payment_proof_url: ['payment_proof_url', 'Payment Capture / 付款截圖', '付款截圖', 'payment_proof'],
    order_status: ['order_status', '訂單狀態', 'status'],
    payment_status: ['payment_status', '付款狀態'],
    order_kind: ['order_kind', '類型'],
    raw_form_row: ['raw_form_row', 'raw'],
    updated_at: ['updated_at', '更新時間'],
  };

  var row = header.map(function () { return ''; });

  Object.keys(props).forEach(function (key) {
    var names = aliases[key] || [key];
    for (var i = 0; i < names.length; i++) {
      var idx = header.indexOf(names[i]);
      if (idx >= 0) {
        row[idx] = props[key];
        return;
      }
    }
  });

  // 若完全對不到 order_id 欄，就整列 append 固定順序（給空白新表）
  var hasAny = row.some(function (c) { return c !== '' && c != null; });
  if (!hasAny) {
    ordersSh.appendRow([
      props.order_id,
      props.created_at,
      props.source,
      props.channel,
      props.customer_name,
      props.email,
      props.phone,
      props.metal_pass_id,
      props.product_summary,
      props.qty_early_bird,
      props.qty_advanced,
      props.total_payable,
      props.payment_proof_url,
      props.order_status,
      props.payment_status,
      props.order_kind,
      props.raw_form_row,
      props.updated_at,
    ]);
    // 若第一列是空的，寫表頭
    if (ordersSh.getLastRow() === 1) {
      ordersSh.insertRowBefore(1);
      ordersSh.getRange(1, 1, 1, 18).setValues([[
        'order_id', 'created_at', 'source', 'channel', 'customer_name', 'email', 'phone',
        'metal_pass_id', 'product_summary', 'qty_early_bird', 'qty_advanced', 'total_payable',
        'payment_proof_url', 'order_status', 'payment_status', 'order_kind', 'raw_form_row', 'updated_at',
      ]]);
    }
    return;
  }

  ordersSh.appendRow(row);
}
