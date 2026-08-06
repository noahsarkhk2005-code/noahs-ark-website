/**
 * 修正 AS_Pie_Tickets / AS_Chart_Tickets 數據
 *
 * 問題：
 * 1) sales_qty 別名（有/無 emoji）被加兩次 → 數量翻倍
 * 2) unit_price 寫入後被 Sheets 當「日期序號」顯示成 1900年…
 * 3) sales_qty 被套用 $ 貨幣格式
 *
 * 正確規則：
 * - label = Order_Categories.tally_column_hint
 * - sales_qty = NA_Tickets 該票種數量加總（每筆訂單只算一次）
 * - unit_price = Order_Categories.list_price（純數字）
 * - revenue = sales_qty * unit_price
 *
 * 執行：fixAppSheetTicketPie
 */

var FIX_ = {
  SRC: 'NA_Tickets',
  CAT: 'Order_Categories',
  PIE: 'AS_Pie_Tickets',
  CHART: 'AS_Chart_Tickets',
  ORDERS: 'AS_TicketOrders'
};

function rg_(sh, r, c, nR, nC) {
  return sh.getRange(r, c, nR, nC);
}

function fixAppSheetTicketPie() {
  var ss = SpreadsheetApp.getActive();
  var catalog = loadTicketCatalog_(ss); // hint -> {label, price}
  var lines = loadTicketLines_(ss);     // [{sid, typeRaw, qty, totalAmount}]

  // 彙總：每個 catalog label 的 qty
  var agg = {}; // label -> qty
  var keys = Object.keys(catalog);
  var i;
  for (i = 0; i < keys.length; i++) {
    agg[catalog[keys[i]].label] = 0;
  }

  var unmatched = [];
  for (i = 0; i < lines.length; i++) {
    var line = lines[i];
    var hit = matchTicketLabel_(line.typeRaw, catalog);
    if (!hit) {
      unmatched.push(line.sid + ': ' + line.typeRaw + ' ×' + line.qty);
      continue;
    }
    agg[hit.label] = (agg[hit.label] || 0) + line.qty;
  }

  // 組列
  var pieRows = [];
  var chartRows = [];
  var orderRows = [];
  var labels = Object.keys(agg);
  // 穩定排序：依 revenue desc
  labels.sort(function (a, b) {
    var pa = catalog[normKey_(a)] ? catalog[normKey_(a)].price : 0;
    var pb = catalog[normKey_(b)] ? catalog[normKey_(b)].price : 0;
    // find price by label
    pa = priceForLabel_(a, catalog);
    pb = priceForLabel_(b, catalog);
    return (agg[b] * pb) - (agg[a] * pa);
  });

  var now = new Date();
  var n = 0;
  for (i = 0; i < labels.length; i++) {
    var lab = labels[i];
    var qty = agg[lab] || 0;
    var price = priceForLabel_(lab, catalog);
    var rev = qty * price;
    n++;
    // AS_Pie_Tickets
    pieRows.push([
      'TKT-' + n,
      'ticket',
      lab,          // tally_column_hint
      lab,          // ticket_type
      lab,          // chart_label
      qty,          // sales_qty  純整數
      rev,          // revenue
      price,        // unit_price 純數字
      now
    ]);
    // AS_Chart_Tickets（極簡）
    if (qty > 0) {
      chartRows.push([
        'CT-' + n,
        lab,
        qty,
        rev,
        'ticket'
      ]);
    }
  }

  // 訂單明細（每筆 NA_Tickets 一列，正確單價）
  for (i = 0; i < lines.length; i++) {
    var L = lines[i];
    var H = matchTicketLabel_(L.typeRaw, catalog);
    var lab2 = H ? H.label : L.typeRaw;
    var pr = H ? H.price : guessPrice_(L.typeRaw);
    // 若 Total Amount 合理且 qty>0，可交叉檢查；仍以 catalog 單價為準
    orderRows.push([
      'TORD-' + L.sid + '-' + (i + 1),
      L.sid,
      L.submittedAt,
      L.name,
      L.email,
      L.phone,
      L.metal,
      lab2,
      L.qty,
      pr,
      L.qty * pr,
      L.proof,
      'ticket',
      now
    ]);
  }

  writePieTicketsFixed_(ss, pieRows);
  writeChartTicketsFixed_(ss, chartRows);
  writeTicketOrdersFixed_(ss, orderRows);

  var summary = [];
  for (i = 0; i < pieRows.length; i++) {
    summary.push(
      pieRows[i][2] + ' → qty=' + pieRows[i][5] +
      ' price=' + pieRows[i][7] +
      ' rev=' + pieRows[i][6]
    );
  }

  SpreadsheetApp.getUi().alert(
    'AS_Pie_Tickets 已修正\n\n' +
    summary.join('\n') +
    (unmatched.length ? '\n\n未對上:\n' + unmatched.join('\n') : '') +
    '\n\n請 AppSheet Sync\nPie 只勾 label + revenue（或 sales_qty）'
  );
}

/* ========== 讀主檔票種 ========== */

function loadTicketCatalog_(ss) {
  // normKey -> {label, price}
  var map = {};
  var sh = ss.getSheetByName(FIX_.CAT);
  if (!sh || sh.getLastRow() < 2) {
    // 後備
    addCat_(map, '會員特級優惠 HKD300', 300);
    addCat_(map, '🎫 早鳥門票 HKD300', 300);
    addCat_(map, '🎫 預售 HKD350', 350);
    return map;
  }
  var lastCol = Math.min(sh.getLastColumn(), 25);
  var lastRow = sh.getLastRow();
  var headers = rg_(sh, 1, 1, 1, lastCol).getValues()[0];
  var idx = headerIndexMap_(headers);
  var data = rg_(sh, 2, 1, lastRow - 1, lastCol).getValues();
  var r;
  for (r = 0; r < data.length; r++) {
    if (idx.level != null && Number(data[r][idx.level]) !== 3) continue;
    var ch = String(data[r][idx.channel] || '').toLowerCase();
    if (ch && ch !== 'ticket') continue;
    var hint = cell_(data[r], idx, 'tally_column_hint');
    var form = cell_(data[r], idx, 'form_option_label');
    var nameZh = cell_(data[r], idx, 'name_zh');
    var label = hint || form || nameZh;
    if (!label) continue;
    var price = toMoney_(idx.list_price != null ? data[r][idx.list_price] : 0);
    if (!price) price = guessPrice_(label);
    addCat_(map, label, price);
    if (hint) addCat_(map, hint, price);
    if (form) addCat_(map, form, price);
    if (nameZh) addCat_(map, nameZh, price);
    addCat_(map, stripEmoji_(label), price);
  }
  return map;
}

function addCat_(map, label, price) {
  if (!label) return;
  var L = String(label).trim();
  var k = normKey_(L);
  if (!k) return;
  // 保留「正式」label：優先有 emoji / 較長的 hint
  if (!map[k] || (L.length >= map[k].label.length)) {
    map[k] = { label: L, price: price };
  }
}

function priceForLabel_(label, catalog) {
  var hit = catalog[normKey_(label)] || catalog[normKey_(stripEmoji_(label))];
  return hit ? hit.price : guessPrice_(label);
}

function matchTicketLabel_(raw, catalog) {
  var s = String(raw || '').trim();
  if (!s) return null;
  var k = normKey_(s);
  if (catalog[k]) return catalog[k];
  var k2 = normKey_(stripEmoji_(s));
  if (catalog[k2]) return catalog[k2];

  // 模糊：catalog key 互相包含
  var keys = Object.keys(catalog);
  var i;
  for (i = 0; i < keys.length; i++) {
    var ck = keys[i];
    if (!ck) continue;
    if (k.indexOf(ck) >= 0 || ck.indexOf(k) >= 0 || k2.indexOf(ck) >= 0 || ck.indexOf(k2) >= 0) {
      return catalog[ck];
    }
  }
  // 關鍵字
  if (/會員|metal\s*pass|特級/i.test(s)) {
    for (i = 0; i < keys.length; i++) {
      if (/會員|metal|特級/i.test(catalog[keys[i]].label)) return catalog[keys[i]];
    }
  }
  if (/早鳥|early/i.test(s)) {
    for (i = 0; i < keys.length; i++) {
      if (/早鳥|early/i.test(catalog[keys[i]].label)) return catalog[keys[i]];
    }
  }
  if (/預售|advanced|presale/i.test(s)) {
    for (i = 0; i < keys.length; i++) {
      if (/預售|advanced|presale/i.test(catalog[keys[i]].label)) return catalog[keys[i]];
    }
  }
  return null;
}

/* ========== 讀 NA_Tickets（每列一次，不重複加） ========== */

function loadTicketLines_(ss) {
  var sh = ss.getSheetByName(FIX_.SRC);
  var out = [];
  if (!sh || sh.getLastRow() < 2) return out;

  var lastCol = Math.min(sh.getLastColumn(), 15);
  var lastRow = sh.getLastRow();
  var headers = rg_(sh, 1, 1, 1, lastCol).getValues()[0];
  var data = rg_(sh, 2, 1, lastRow - 1, lastCol).getValues();
  var h = headerIndexMap_(headers);

  var cSid = firstCol_(h, ['Submission ID']);
  var cAt = firstCol_(h, ['Submitted at']);
  var cName = firstCol_(h, ['💀 Name/稱呼', 'Name/稱呼', 'Name']);
  var cEmail = firstCol_(h, ['📩 Email/電郵', 'Email']);
  var cTel = firstCol_(h, ['☎️ Tel/電話號碼', 'Tel']);
  var cPass = firstCol_(h, ['Metal Pass 會員編號', 'Metal Pass']);
  var cType = firstCol_(h, ['門票類別', '票種']);
  var cQty = firstCol_(h, ['數量', 'Qty']);
  var cProof = firstCol_(h, ['Payment Capture', '付款截圖']);

  var r;
  if (cType != null && cQty != null) {
    for (r = 0; r < data.length; r++) {
      var type = String(data[r][cType] || '').trim();
      var qty = toQty_(data[r][cQty]);
      if (!type || qty <= 0) continue;
      var sid = cSid != null ? String(data[r][cSid] || '').trim() : ('R' + (r + 2));
      // 跳過明顯測試可選：if (sid.indexOf('TEST-')===0) continue;
      out.push({
        sid: sid || ('R' + (r + 2)),
        submittedAt: cAt != null ? data[r][cAt] : '',
        name: cName != null ? data[r][cName] : '',
        email: cEmail != null ? data[r][cEmail] : '',
        phone: cTel != null ? data[r][cTel] : '',
        metal: cPass != null ? data[r][cPass] : '',
        typeRaw: type,
        qty: qty,
        proof: cProof != null ? data[r][cProof] : ''
      });
    }
    return out;
  }

  // 舊多欄格式
  var c;
  for (c = 0; c < headers.length; c++) {
    var hh = String(headers[c] || '').trim();
    if (!/早鳥|預售|會員|Metal|Early|Advanced|門票/i.test(hh)) continue;
    for (r = 0; r < data.length; r++) {
      var q = toQty_(data[r][c]);
      if (q <= 0) continue;
      var sid2 = cSid != null ? String(data[r][cSid] || '').trim() : ('R' + (r + 2));
      out.push({
        sid: sid2 || ('R' + (r + 2)),
        submittedAt: cAt != null ? data[r][cAt] : '',
        name: cName != null ? data[r][cName] : '',
        email: cEmail != null ? data[r][cEmail] : '',
        phone: cTel != null ? data[r][cTel] : '',
        metal: cPass != null ? data[r][cPass] : '',
        typeRaw: hh,
        qty: q,
        proof: cProof != null ? data[r][cProof] : ''
      });
    }
  }
  return out;
}

/* ========== 寫入（強制數字格式，避免變成日期） ========== */

function writePieTicketsFixed_(ss, rows) {
  var headers = [
    'row_id', 'channel', 'tally_column_hint', 'ticket_type', 'chart_label',
    'sales_qty', 'revenue', 'unit_price', 'updated_at'
  ];
  writeNumTable_(ss, FIX_.PIE, headers, rows, '#fce8e6', {
    qty: 6, rev: 7, price: 8, date: 9
  });
}

function writeChartTicketsFixed_(ss, rows) {
  var headers = ['chart_key', 'label', 'sales_qty', 'revenue', 'channel'];
  writeNumTable_(ss, FIX_.CHART, headers, rows, '#f4cccc', {
    qty: 3, rev: 4, price: 0, date: 0
  });
}

function writeTicketOrdersFixed_(ss, rows) {
  var headers = [
    'order_key', 'submission_id', 'submitted_at', 'customer_name', 'email', 'phone',
    'metal_pass', 'ticket_label', 'sales_qty', 'unit_price', 'revenue',
    'payment_proof', 'channel', 'updated_at'
  ];
  writeNumTable_(ss, FIX_.ORDERS, headers, rows, '#fce8e6', {
    qty: 9, rev: 11, price: 10, date: 14
  });
}

/**
 * colMap: 1-based column numbers for formats
 */
function writeNumTable_(ss, name, headers, body, color, colMap) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  var cols = headers.length;
  var nBody = body.length;
  var total = 1 + nBody;

  if (total > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), total - sh.getMaxRows() + 3);
  if (cols > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns() + 2);

  var clearR = Math.min(Math.max(sh.getLastRow(), total), 200);
  var clearC = Math.min(Math.max(sh.getLastColumn(), cols), 20);
  try {
    rg_(sh, 1, 1, clearR, clearC).clearContent();
    // 清掉可能把數字變日期的格式
    rg_(sh, 1, 1, clearR, clearC).setNumberFormat('@'); // 先當文字清格式
  } catch (e) {}

  var all = [headers];
  var i;
  for (i = 0; i < nBody; i++) all.push(body[i]);
  rg_(sh, 1, 1, all.length, cols).setValues(all);

  rg_(sh, 1, 1, 1, cols).setFontWeight('bold').setBackground(color);
  sh.setFrozenRows(1);

  if (nBody > 0) {
    // 強制數字欄為數字（不是日期、不是貨幣字串）
    if (colMap.qty > 0) {
      rg_(sh, 2, colMap.qty, nBody, 1).setNumberFormat('0');
    }
    if (colMap.rev > 0) {
      rg_(sh, 2, colMap.rev, nBody, 1).setNumberFormat('0');
    }
    if (colMap.price > 0) {
      rg_(sh, 2, colMap.price, nBody, 1).setNumberFormat('0');
    }
    if (colMap.date > 0) {
      rg_(sh, 2, colMap.date, nBody, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    }
  }
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

function toQty_(v) {
  if (v === '' || v == null) return 0;
  var n = Number(v);
  return isFinite(n) ? Math.max(0, Math.floor(n)) : 0;
}

function toMoney_(v) {
  if (v === '' || v == null) return 0;
  if (typeof v === 'number') {
    // 若被當日期序號（很大或奇怪）不在此處理；list_price 應是 80–350
    return isFinite(v) ? v : 0;
  }
  var n = Number(String(v).replace(/[^0-9.\-]/g, ''));
  return isFinite(n) ? n : 0;
}

function guessPrice_(h) {
  var m = String(h || '').match(/HKD?\s*(\d+)/i);
  return m ? Number(m[1]) : 0;
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
