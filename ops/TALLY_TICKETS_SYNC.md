# Tally 購票 → Sheet → AppSheet（演出票物）

## 資料流

```
Tally https://tally.so/r/J9b2zY
   ↓ Integrations → Google Sheets
分頁：NA_Tickets（你現在的表頭）
   ↓ Apps Script「同步 Tally 購票 → Orders」
Orders（source = ticket_form，狀態 = 待審核）
   ↓ AppSheet 同步 / 挪亞審核
演出票物（訂單）
   ↓ 審核「通過／確認付款」
AS_Tickets（出票 + QR）
```

## 你的 Tally 表頭（NA_Tickets）

| 欄位 | 用途 |
|------|------|
| Submission ID | 去重主鍵（必填） |
| Respondent ID | 作廢參考 |
| Submitted at | 下單時間 |
| Total Amount | 應付總價（審核對截圖） |
| 💀 Name/稱呼 | 顧客姓名 |
| 📩 Email/電郵 | 電郵 |
| ☎️ Tel/電話號碼 | 電話 |
| Metal Pass 會員編號 | 會員折扣／綁定 |
| Untitled number field | **建議刪除或改名**（未命名欄，易錯） |
| 🎫 早鳥門票 HKD300 | 數量 → Early Bird |
| 🎫 預售 HKD350 | 數量 → Advanced |
| Payment Capture / 付款截圖 | 截圖 URL |

## 必改 Settings（試算表 Settings 分頁）

| key | 新 value |
|-----|----------|
| `TICKET_SOURCE_TAB_NAME` | `NA_Tickets`（你實際分頁名） |
| `TICKET_SOURCE_SPREADSHEET_ID` | `1DgB01yyo8cdSB6IdsIv2lzYtYGlNkPd1WwclgQc07Bg` |

舊值 `挪亞方舟2.5 - PROTOSS 20週年` 已失效，不改則同步不到新單。

## 必改 Products 價錢（與 Tally 一致）

| sku | list_price | form_option_label 建議 |
|-----|------------|------------------------|
| T-EB-280（可改名 T-EB-300） | **300** | 早鳥 / Early Bird / 🎫 早鳥門票 HKD300 |
| T-ADV-380（可改名 T-ADV-350） | **350** | 預售 / Advanced / 🎫 預售 HKD350 |
| T-METAL-280 | 280 | Metal Pass（若表單有） |

## 安裝同步腳本

1. 試算表 → **擴充功能** → **Apps Script**
2. 新增檔案，貼上 `sync_tally_tickets.gs` 內容
3. 儲存 → 重新整理試算表
4. 選單 **NOAHSARK** → **同步 Tally 購票 → Orders**
5. 填一筆測試單 → 再跑同步 → 檢查 `Orders` / AppSheet

可選：觸發條件 → 時間驅動 → 每 5 分鐘執行 `syncTallyTicketsToOrders`。

## AppSheet

1. Data → Sync
2. 演出票物／訂單讀 **AS_Orders**（或 Orders 對應表）
3. 新單 `order_kind=購票`、`status=待審核`
4. **通過後**才會寫入 **AS_Tickets**（QR 票物）
