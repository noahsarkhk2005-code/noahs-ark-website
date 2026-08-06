# AppSheet 圓餅圖資料庫（NA_Tickets + NA_Merch）

## 資料流

```
NA_Tickets  ──┐
              ├─► buildAppSheetPieData ──► AS_Pie_Tickets / AS_Pie_Merch / AS_Pie_All
NA_Merch    ──┘         ↑
                 Order_Categories（單價）
```

## 三張表（給 AppSheet）

### 1. `AS_Pie_Tickets` — 票務（按**門票類型**）

| 欄位 | 用途 |
|------|------|
| `ticket_type` / `chart_label` | 圓餅分類（早鳥／預售／會員…） |
| `sales_qty` | 總銷量 |
| `revenue` | 總銷售額（數量×單價） |

**Pie 建議**
- 銷售額：Category = `chart_label`，Value = `revenue`
- 銷量：Category = `chart_label`，Value = `sales_qty`

### 2. `AS_Pie_Merch` — 商品（按**單個貨品／分欄**）

| 欄位 | 用途 |
|------|------|
| `product_label` / `chart_label` | 圓餅分類（Tee／毛巾／服飾…） |
| `sales_qty` | 該貨品總銷量 |
| `revenue` | 該貨品總銷售額 |

**Pie 建議**
- 銷售額：Category = `chart_label`，Value = `revenue`
- 銷量：Category = `chart_label`，Value = `sales_qty`

### 3. `AS_Pie_All` — 合併（票務+商品）

| 欄位 | 用途 |
|------|------|
| `channel_zh` | 票務 / 商品（可做 filter） |
| `item_label` / `chart_label` | 項目名稱 |
| `sales_qty` / `revenue` | 銷量／銷售額 |

可用 Slice：`[channel]="ticket"` 或 `merch`。

## 更新資料

1. Apps Script 貼上 `build_appsheet_pie_data.gs`
2. 執行 **`buildAppSheetPieData`**
3. AppSheet → **Data** 加入三表 → 主鍵 `row_id` → **Sync**
4. 有新 Tally 單後再執行一次腳本

## AppSheet 圖表設定（簡）

1. UX → Chart view  
2. Chart type = **Pie**  
3. Source = `AS_Pie_Tickets` 或 `AS_Pie_Merch`  
4. Label column = `chart_label`  
5. Value column = `revenue`（或 `sales_qty`）  
6. 做兩張圖：一張銷售額、一張銷量  

## 主鍵建議

| 表 | Key |
|----|-----|
| AS_Pie_Tickets | `row_id` |
| AS_Pie_Merch | `row_id` |
| AS_Pie_All | `row_id` |

Updates: **Read-only**（由腳本覆寫即可）
