/**
 * 將 Order_Categories 寫成「所有銷售」主檔
 * 大類：票務 / 商品
 * 商品細類：服飾 / 配件 / 毛巾 / Pack / 袋類
 * SKU 列：對應 Tally 欄位 + 價錢
 *
 * 用法：Apps Script 新增檔案 → 貼上 → 執行 seedOrderCategories
 * 或選單 NOAHSARK → 初始化 Order_Categories 銷售主檔
 */

var OC_HEADERS_ = [
  'category_id', 'parent_id', 'level', 'channel', 'sub_category', 'sub_category_zh',
  'sku', 'name_zh', 'name_en', 'list_price', '銷售數量', '單件貨品總收入', 'form_option_label',
  'tally_source_tab', 'tally_column_hint', 'track_inventory', 'is_active',
  'sort_order', 'appsheet_group', 'notes'
];

// J=list_price K=銷售數量 L=單件貨品總收入（數量×售價，由 refresh 重算）
var OC_ROWS_ = [
  // 大類
  ['CAT-TICKET', '', 1, 'ticket', '', '', '', '票務', 'Tickets', '', '', '', '', '', '', '', true, 1, '票務', '大類：現場門票'],
  ['CAT-MERCH', '', 1, 'merch', '', '', '', '商品', 'Merch', '', '', '', '', '', '', '', true, 10, '商品', '大類：周邊商品'],
  // 票務細類
  ['CAT-TICKET-MEMBER', 'CAT-TICKET', 2, 'ticket', 'member', '會員票種', '', '', '', '', '', '', '', '', '', '', true, 2, '票務', '票務細類'],
  ['CAT-TICKET-EARLY', 'CAT-TICKET', 2, 'ticket', 'early_bird', '早鳥', '', '', '', '', '', '', '', '', '', '', true, 3, '票務', '票務細類'],
  ['CAT-TICKET-ADV', 'CAT-TICKET', 2, 'ticket', 'advanced', '預售', '', '', '', '', '', '', '', '', '', '', true, 4, '票務', '票務細類'],
  // 商品細類
  ['CAT-MERCH-APPAREL', 'CAT-MERCH', 2, 'merch', 'apparel', '服飾', '', '', '', '', '', '', '', '', '', '', true, 11, '商品', '商品細類'],
  ['CAT-MERCH-ACCESSORY', 'CAT-MERCH', 2, 'merch', 'accessory', '配件', '', '', '', '', '', '', '', '', '', '', true, 12, '商品', '商品細類'],
  ['CAT-MERCH-TOWEL', 'CAT-MERCH', 2, 'merch', 'towel', '毛巾', '', '', '', '', '', '', '', '', '', '', true, 13, '商品', '商品細類'],
  ['CAT-MERCH-PACK', 'CAT-MERCH', 2, 'merch', 'pack', '套裝 Pack', '', '', '', '', '', '', '', '', '', '', true, 14, '商品', '商品細類'],
  ['CAT-MERCH-BAG', 'CAT-MERCH', 2, 'merch', 'bag', '袋類', '', '', '', '', '', '', '', '', '', '', true, 15, '商品', '商品細類'],
  // SKU — list_price, 銷售數量=0, 單件貨品總收入=0
  ['SKU-T-METAL-300', 'CAT-TICKET-MEMBER', 3, 'ticket', 'member', '會員票種', 'T-METAL-300', '會員特級優惠', 'Member Special', 300, 0, 0, '會員特級優惠 HKD300', 'NA_Tickets', '會員特級優惠 HKD300', false, true, 101, '票務', 'Tally 購票'],
  ['SKU-T-EB-300', 'CAT-TICKET-EARLY', 3, 'ticket', 'early_bird', '早鳥', 'T-EB-300', '早鳥門票', 'Early Bird', 300, 0, 0, '🎫 早鳥門票 HKD300', 'NA_Tickets', '🎫 早鳥門票 HKD300', false, true, 102, '票務', 'Tally 購票'],
  ['SKU-T-ADV-350', 'CAT-TICKET-ADV', 3, 'ticket', 'advanced', '預售', 'T-ADV-350', '預售門票', 'Presale', 350, 0, 0, '🎫 預售 HKD350', 'NA_Tickets', '🎫 預售 HKD350', false, true, 103, '票務', 'Tally 購票'],
  ['SKU-M-TEE-BLK-M', 'CAT-MERCH-APPAREL', 3, 'merch', 'apparel', '服飾', 'M-TEE-BLK-M', '活動 Tee 黑 M', 'Event Tee Black M', 280, 0, 0, 'Ark T-Shirt HKD280', 'NA_Merch', 'Ark T-Shirt HKD280', true, true, 201, '商品', 'Tally 商品'],
  ['SKU-M-ARK-TOTE', 'CAT-MERCH-BAG', 3, 'merch', 'bag', '袋類', 'M-ARK-TOTE', 'Ark Tote Bag', 'Ark Tote Bag', 120, 0, 0, 'Ark Tote Bag HKD120', 'NA_Merch', 'Ark Tote Bag HKD120', true, true, 202, '商品', 'Tally 商品'],
  ['SKU-M-ARK-KEYCHAIN', 'CAT-MERCH-ACCESSORY', 3, 'merch', 'accessory', '配件', 'M-ARK-KEYCHAIN', 'Ark Keychain', 'Ark Keychain', 120, 0, 0, 'Ark Keychain/鎖匙扣 HKD120', 'NA_Merch', 'Ark Keychain/鎖匙扣 HKD120', true, true, 203, '商品', 'Tally 商品'],
  ['SKU-M-PATCH', 'CAT-MERCH-ACCESSORY', 3, 'merch', 'accessory', '配件', 'M-PATCH', '布章', 'Patch', 100, 0, 0, 'Patch', 'NA_Merch', '', true, true, 204, '商品', '可選'],
  ['SKU-M-ARK-MOUSEPAD', 'CAT-MERCH-ACCESSORY', 3, 'merch', 'accessory', '配件', 'M-ARK-MOUSEPAD', 'Ark Mousepad', 'Ark Mousepad', 40, 0, 0, 'Ark Mousepad', 'NA_Merch', '', true, true, 205, '商品', '可選'],
  ['SKU-M-ARK-TOWER', 'CAT-MERCH-TOWEL', 3, 'merch', 'towel', '毛巾', 'M-ARK-TOWER', 'Ark Tower 毛巾', 'Ark Tower', 80, 0, 0, 'Ark Tower/毛巾 HKD80', 'NA_Merch', 'Ark Tower/毛巾 HKD80', true, true, 211, '商品', 'Tally 商品'],
  ['SKU-M-ARK-BIGPACK', 'CAT-MERCH-PACK', 3, 'merch', 'pack', '套裝 Pack', 'M-ARK-BIGPACK', 'Ark BigPack', 'Ark BigPack', 200, 0, 0, 'Ark BigPack HKD200', 'NA_Merch', 'Ark BigPack HKD200', true, true, 221, '商品', 'Tally 商品'],
  ['SKU-M-ARK-TINYPACK', 'CAT-MERCH-PACK', 3, 'merch', 'pack', '套裝 Pack', 'M-ARK-TINYPACK', 'Ark TinyPack', 'Ark TinyPack', 60, 0, 0, 'Ark TinyPack HKD60', 'NA_Merch', 'Ark TinyPack HKD60', true, true, 222, '商品', 'Tally 商品']
];

function seedOrderCategories() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Order_Categories');
  if (!sh) sh = ss.insertSheet('Order_Categories');

  sh.clear();
  sh.getRange(1, 1, 1, OC_HEADERS_.length).setValues([OC_HEADERS_]);
  sh.getRange(1, 1, 1, OC_HEADERS_.length).setFontWeight('bold');
  sh.setFrozenRows(1);

  sh.getRange(2, 1, OC_ROWS_.length, OC_HEADERS_.length).setValues(OC_ROWS_);
  sh.autoResizeColumns(1, Math.min(8, OC_HEADERS_.length));

  // Settings 標記：銷售主檔分頁
  upsertSetting_(ss, 'SALES_MASTER_TAB', 'Order_Categories');
  upsertSetting_(ss, 'SALES_MASTER_NOTE', '所有可售品項以 Order_Categories 為準；level=3 為 SKU');

  SpreadsheetApp.getUi().alert(
    'Order_Categories 已寫入\n' +
    '大類：票務 / 商品\n' +
    'SKU 列數：' + countSkuRows_() + '\n' +
    'Settings.SALES_MASTER_TAB = Order_Categories'
  );
}

function countSkuRows_() {
  var n = 0;
  var i;
  for (i = 0; i < OC_ROWS_.length; i++) {
    if (OC_ROWS_[i][2] === 3) n++;
  }
  return n;
}

function upsertSetting_(ss, key, value) {
  var sh = ss.getSheetByName('Settings');
  if (!sh) return;
  var vals = sh.getDataRange().getValues();
  var i;
  for (i = 1; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

/** 給 AppSheet / 腳本：只取啟用中的 SKU */
function getActiveSalesSkus() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Order_Categories');
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  var header = data[0];
  var idx = {};
  var c;
  for (c = 0; c < header.length; c++) idx[String(header[c])] = c;

  var out = [];
  var r;
  for (r = 1; r < data.length; r++) {
    var row = data[r];
    if (Number(row[idx.level]) !== 3) continue;
    if (String(row[idx.is_active]).toUpperCase() === 'FALSE') continue;
    out.push({
      category_id: row[idx.category_id],
      channel: row[idx.channel],
      sub_category: row[idx.sub_category],
      sub_category_zh: row[idx.sub_category_zh],
      sku: row[idx.sku],
      name_zh: row[idx.name_zh],
      name_en: row[idx.name_en],
      list_price: row[idx.list_price],
      sales_qty: (idx['銷售數量'] != null) ? row[idx['銷售數量']] : row[idx.member_price],
      form_option_label: row[idx.form_option_label],
      tally_source_tab: row[idx.tally_source_tab],
      tally_column_hint: row[idx.tally_column_hint],
      track_inventory: row[idx.track_inventory],
      appsheet_group: row[idx.appsheet_group]
    });
  }
  return out;
}

/** 可選：把 menu 加到現有 onOpen（若已有 onOpen 請手動合併） */
function addOrderCategoriesMenu() {
  SpreadsheetApp.getUi()
    .createMenu('NOAHSARK-銷售主檔')
    .addItem('初始化 Order_Categories 銷售主檔', 'seedOrderCategories')
    .addToUi();
}
