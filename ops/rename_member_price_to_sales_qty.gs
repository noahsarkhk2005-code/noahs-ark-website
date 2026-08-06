/**
 * Order_Categories：把 member_price 欄改名為「銷售數量」
 * - 只改表頭，不刪欄
 * - SKU 列（level=3）若原是會員價數字，重設為 0（銷售數量起點）
 * - 大類/細類列清空該欄
 *
 * 執行：renameMemberPriceToSalesQty
 */

function renameMemberPriceToSalesQty() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Order_Categories');
  if (!sh) throw new Error('找不到 Order_Categories');

  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  if (lastCol < 1 || lastRow < 1) throw new Error('Order_Categories 是空的');

  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var col = -1;
  var c;
  for (c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim();
    if (h === 'member_price' || h === 'member price' || h === '會員價' || h === '銷售數量') {
      col = c;
      break;
    }
  }
  if (col < 0) {
    // 常見位置：list_price 後面 = 第 11 欄 (index 10)
    if (headers.length >= 11) {
      col = 10;
    } else {
      throw new Error('找不到 member_price 欄。目前表頭: ' + headers.join(' | '));
    }
  }

  // 1) 改表頭
  sh.getRange(1, col + 1).setValue('銷售數量');
  sh.getRange(1, col + 1).setFontWeight('bold').setBackground('#fff2cc');

  // 2) 找 level 欄
  var levelCol = -1;
  for (c = 0; c < headers.length; c++) {
    if (String(headers[c] || '').trim() === 'level') {
      levelCol = c;
      break;
    }
  }

  var resetSku = 0;
  var cleared = 0;
  if (lastRow >= 2) {
    var numRows = lastRow - 1;
    var levels = levelCol >= 0
      ? sh.getRange(2, levelCol + 1, numRows, 1).getValues()
      : null;
    var vals = sh.getRange(2, col + 1, numRows, 1).getValues();
    var r;
    for (r = 0; r < numRows; r++) {
      var level = levels ? Number(levels[r][0]) : 3;
      if (level === 3) {
        // 舊會員價 → 銷售數量從 0 起（避免把 255/298 當已售量）
        vals[r][0] = 0;
        resetSku++;
      } else {
        vals[r][0] = '';
        cleared++;
      }
    }
    sh.getRange(2, col + 1, numRows, 1).setValues(vals);
  }

  SpreadsheetApp.getUi().alert(
    '完成\n' +
    '表頭：member_price → 銷售數量（第 ' + (col + 1) + ' 欄）\n' +
    'SKU 列重設為 0：' + resetSku + '\n' +
    '非 SKU 列清空：' + cleared + '\n\n' +
    'AppSheet 請重新 Sync，欄位名稱改為「銷售數量」。'
  );
}
