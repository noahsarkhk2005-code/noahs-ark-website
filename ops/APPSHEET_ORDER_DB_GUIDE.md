# AppSheet 票務／商品訂單資料庫（Pie 專用格式）

## 為什麼不要勾 `_RowNumber`

AppSheet 圖表若 **Chart columns 全選**，會出現 `_RowNumber`，Pie 無法正確分類。  
圖表專用表只有 **5 個欄位**，請**只勾 2 個**。

## 四張表

| 分頁 | 用途 | 主鍵 |
|------|------|------|
| **AS_TicketOrders** | 票務訂單明細（一單一列） | `order_key` |
| **AS_MerchOrders** | 商品訂單明細 | `order_key` |
| **AS_Chart_Tickets** | 票務 Pie（極簡） | `chart_key` |
| **AS_Chart_Merch** | 商品 Pie（極簡） | `chart_key` |

## 圖表表結構（重要）

### AS_Chart_Tickets / AS_Chart_Merch

| 欄位 | 類型 | 說明 |
|------|------|------|
| `chart_key` | Text · **KEY** | 主鍵 |
| `label` | Text | **分類名稱 = tally_column_hint** |
| `sales_qty` | **Number** | 總銷量 |
| `revenue` | **Number** | 總銷售額 |
| `channel` | Text | ticket / merch |

### Pie 設定（AppSheet）

1. View type = **Chart** → Pie  
2. Data = `AS_Chart_Tickets` 或 `AS_Chart_Merch`  
3. **Chart columns** 只勾：  
   - ✅ `label`  
   - ✅ `revenue`　← 看銷售額  
   - 或 ✅ `label` + ✅ `sales_qty`　← 看銷量  
4. ❌ 不要勾 `_RowNumber`  
5. ❌ 不要勾 `chart_key` / `channel`（可選勾 channel 做 filter）

建議做 **4 張圖**：

| 圖 | 表 | columns |
|----|-----|--------|
| 票務·銷售額 | AS_Chart_Tickets | revenue |
| 票務·銷量 | AS_Chart_Tickets | sales_qty |
| 商品·銷售額 | AS_Chart_Merch | revenue |
| 商品·銷量 | AS_Chart_Merch | sales_qty |

## 訂單明細結構

### AS_TicketOrders

`order_key | submission_id | submitted_at | customer_name | email | phone | metal_pass | ticket_label | sales_qty | unit_price | revenue | payment_proof | channel | updated_at`

- `ticket_label` = Order_Categories.**tally_column_hint**

### AS_MerchOrders

同上，`product_label` = **tally_column_hint**

## 產生／更新資料

1. Apps Script 貼上 `build_appsheet_order_db.gs`  
2. 執行 **`buildAppSheetOrderDb`**  
3. AppSheet → Data 加入四表  
4. 每表設定：  
   - Key = `order_key` 或 `chart_key`  
   - `sales_qty` / `revenue` 類型 = **Number**  
   - Updates = **Read-only**（建議）  
5. **Regenerate / Sync**

## 資料流

```
NA_Tickets / NA_Merch
        +
Order_Categories.tally_column_hint + list_price
        ↓
AS_TicketOrders / AS_MerchOrders   ← 訂單明細
        ↓ 彙總
AS_Chart_Tickets / AS_Chart_Merch ← Pie 只用這兩張
```
