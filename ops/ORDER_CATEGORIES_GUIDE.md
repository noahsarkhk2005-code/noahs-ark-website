# Order_Categories — 所有銷售主檔

**分頁：** `Order_Categories`（gid 以試算表為準，你連結的是此分頁）  
**試算表：** [NoahArk_Event_Management_2026](https://docs.google.com/spreadsheets/d/1DgB01yyo8cdSB6IdsIv2lzYtYGlNkPd1WwclgQc07Bg/edit?gid=411439402)

## 定位

此分頁是 **所有可售品項的唯一資料來源（Source of Truth）**：

| 用途 | 怎麼用 |
|------|--------|
| 票務 / 商品分流 | `channel` = `ticket` \| `merch` |
| 商品細分類 | `sub_category` / `sub_category_zh` |
| 價錢 | `list_price` |
| 銷售數量 | **K 欄** `銷售數量` ← Tally 訂單彙總（`NA_Tickets` + `NA_Merch`） |
| Tally 對欄 | `tally_source_tab` + `tally_column_hint` / `form_option_label` |
| AppSheet 分組 | `appsheet_group` = 票務 / 商品 |
| 庫存 | `track_inventory` + 對應 `Products` / `Inventory` 的 `sku` |

Tally 原始提交仍寫入 `NA_Tickets` / `NA_Merch`；**同步進 Orders 時應以本表 SKU 為準**（價錢、名稱、分類）。

## 三層結構

```
level 1  大類     CAT-TICKET 票務 | CAT-MERCH 商品
level 2  細類     會員票種 / 早鳥 / 預售
                  服飾 / 配件 / 毛巾 / Pack / 袋類
level 3  SKU      真正可賣的品項（有 sku + 價錢）
```

## 初始化（建議）

1. Apps Script 新增檔案，貼上 `seed_order_categories.gs`
2. 執行 `seedOrderCategories`（或選單「初始化 Order_Categories 銷售主檔」）
3. 檢查 `Order_Categories` 是否有表頭 + 資料
4. Settings 會寫入：`SALES_MASTER_TAB = Order_Categories`

也可直接開 `Order_Categories.csv` → 全選貼到分頁 A1。

## 欄位說明

| 欄位 | 說明 |
|------|------|
| category_id | 主鍵（CAT- / SKU-） |
| parent_id | 上層 category_id |
| level | 1 大類 / 2 細類 / 3 SKU |
| channel | ticket 或 merch |
| sub_category | 英文細類 key |
| sub_category_zh | 中文細類 |
| sku | 對齊 Products.sku |
| list_price | 售價 |
| 銷售數量（**K 欄**） | 已售數量 = Σ Tally 訂單該品項數量；腳本固定寫入 **K 欄** |

### 銷售數量怎麼算

```
NA_Tickets：門票類別 + 數量  → 對 form_option_label / tally_column_hint
NA_Merch  ：商品分欄或各商品欄 + 數量 → 同上
         ↓
Order_Categories.銷售數量（每個 SKU）
```

執行：Apps Script → **`refreshSalesQtyFromTally`**  
（虛擬測試單 `TEST-*` 也會計入，方便驗證；正式統計可先 `clearTallyVirtualOrders`）
| form_option_label | Tally 選項全文 |
| tally_source_tab | NA_Tickets 或 NA_Merch |
| tally_column_hint | Tally 寫入 Sheet 的欄名 |
| is_active | FALSE = 停售 |
| appsheet_group | AppSheet 顯示分組：票務 / 商品 |

## 與其他分頁關係

```
Order_Categories (主檔：可賣什麼)
       │
       ├─► Products   （可同步 sku 價錢；或以本表取代）
       ├─► Inventory  （sku 庫存）
       ├─► NA_Tickets / NA_Merch（Tally 原始單）
       └─► Orders / 挪亞審核 / AppSheet 演出票物・商品
```

## AppSheet 建議

1. Data 加入 `Order_Categories`，主鍵 `category_id`
2. Slice **可售 SKU**：`AND([level]=3, [is_active]=TRUE)`
3. Slice **票務 SKU**：`AND([level]=3, [channel]="ticket", [is_active]=TRUE)`
4. Slice **商品 SKU**：`AND([level]=3, [channel]="merch", [is_active]=TRUE)`
5. 訂單明細的 `sku` Ref → `Order_Categories[sku]`（或 ref category_id）

## 新增商品流程

1. 在 `Order_Categories` 加一列 level=3  
2. 填 `parent_id`（對應細類）、`sku`、價錢、`form_option_label`、`tally_*`  
3. `is_active=TRUE`  
4. 若需庫存：`track_inventory=TRUE` 並在 Inventory 加 sku  
5. Tally 表單選項文字 = `form_option_label`  
6. AppSheet Sync  
