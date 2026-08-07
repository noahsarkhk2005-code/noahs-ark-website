# Noah's Ark 2.0 — 手機訂單管理（AppSheet）

依現有 **Tally → NA_Tickets / NA_Merch → Order_Categories** 格式，重建一套可在手機操作的 2.0 系統。

## 功能一覽

| 功能 | 怎麼做 |
|------|--------|
| 人手管理訂單 | `v2_Orders` 列表 + 詳情 |
| 顯示付款憑證 | `payment_proof_url` 設 **Image** |
| 審核通過／拒絕 | 寫入 `as_action` = `通過` / `拒絕` |
| QR 出票 | 通過票務單自動出票 → `v2_Tickets` |
| 掃碼入場 | 掃 `qr_payload` 或開 Web App |
| 寄郵件 | `as_action` = `寄出郵件` |
| 分析數據 | `v2_Stats` 儀表 + `v2_Chart` 圓餅 |

---

## 一、Google 試算表安裝（5 分鐘）

1. 打開 [NoahArk_Event_Management_2026](https://docs.google.com/spreadsheets/d/1DgB01yyo8cdSB6IdsIv2lzYtYGlNkPd1WwclgQc07Bg/edit)
2. **擴充功能 → Apps Script**
3. 新增檔案，貼上 `ops/v2/noah_ark_v2.gs` 全文 → 儲存
4. 執行 **`installNoahArkV2`**（首次授權）
5. 重新整理試算表，應見選單 **NOAHSARK-V2** 與分頁：

| 分頁 | 用途 |
|------|------|
| `v2_Orders` | 全部訂單 |
| `v2_Tickets` | 入場票 + QR 圖 |
| `v2_ScanLog` | 入場紀錄 |
| `v2_Stats` | 儀表數字 |
| `v2_Products` | 品項主檔 |
| `v2_Chart` | Pie 圖資料 |

6. 部署掃碼 Web App（選單「顯示掃碼 Web App 說明」）

### 日常腳本

| 選單 | 作用 |
|------|------|
| 同步 Tally → v2_Orders | 讀 NA_Tickets / NA_Merch 新單 |
| 處理 as_action | 執行通過／拒絕／出票／寄信 |
| 刷新統計 + 圖表 | 更新 v2_Stats / v2_Chart |
| 一鍵 | 以上三步 |

建議：觸發條件 → 時間驅動 → 每 5–10 分鐘跑 `v2DailyRun`。

---

## 二、AppSheet 新建 App（手機）

1. https://www.appsheet.com → **Create → App → Google Sheets**
2. 選同一試算表
3. 只勾選（或之後加）：

```
v2_Orders
v2_Tickets
v2_ScanLog
v2_Stats
v2_Products
v2_Chart
```

### 主鍵與類型（必設）

| 表 | Key | 重要欄位類型 |
|----|-----|----------------|
| v2_Orders | `order_id` | `payment_proof_url`=**Image**；qty/price/revenue=**Number** |
| v2_Tickets | `ticket_id` | `qr_image_url`=**Image**；`qr_payload`=Text；status=Text |
| v2_ScanLog | `scan_id` | |
| v2_Stats | `metric` | `value`=Number |
| v2_Products | `sku` | |
| v2_Chart | `chart_key` | `sales_qty`/`revenue`=Number |

Updates 建議：

- `v2_Orders`：Updates only（可改 `as_action`、`notes_admin`）
- `v2_Tickets`：Updates only（可改 `as_action=入場`）
- 其餘：Read-only

---

## 三、手機畫面（UX）建議

### 主畫面底部 Tab

1. **待審核** — Slice `v2_Orders`：`[order_status]="待審核"`
2. **全部訂單** — `v2_Orders`
3. **入場掃碼** — Form on `v2_ScanLog` 或掃 `v2_Tickets`
4. **儀表板** — `v2_Stats` + Chart on `v2_Chart`
5. **已入場** — Slice `v2_Tickets`：`[status]="已入場"`

### 訂單詳情（v2_Orders Detail）

顯示：

- `list_title`
- 顧客、品項、數量、金額
- **`payment_proof_url`（大圖）**
- 狀態 badges
- 動作按鈕（見下）

### 動作按鈕（as_action 模式）

在 AppSheet 建 **Actions**（Data: set the values of some columns）：

| 按鈕 | 寫入 |
|------|------|
| ✅ 通過 | `as_action` = `通過` |
| ❌ 拒絕 | `as_action` = `拒絕` |
| 🎫 出票 | `as_action` = `出票` |
| ✉️ 寄信 | `as_action` = `寄出郵件` |

然後用 **Bot / Webhook / Apps Script 時間觸發** 跑 `processV2Actions`  
（或人手按試算表選單「處理 as_action」）。

> 通過票務單會自動出票，寫入 `v2_Tickets`（含 QR 圖）。

### 入場票詳情（v2_Tickets）

- 大圖 `qr_image_url`
- `qr_payload` 文字
- 狀態：有效 / 已入場 / 作廢
- 按鈕「入場」→ `as_action`=`入場` → 跑 `processV2TicketGateActions`  
  或現場掃碼呼叫 Web App

### 掃碼入場（兩種）

**A. Web App（推薦現場）**

```
GET {V2_SCAN_WEBAPP_URL}?payload={qr_payload}&by={掃碼人}
```

回傳 JSON：`result=OK|FAIL`，`result_title`，`message`，`customer_name`

**B. AppSheet 相機掃碼**

- 用「Scan barcode」寫入 `v2_ScanLog.payload` 後 Webhook 到 Web App  
- 或掃到字串後 App 開啟 Web App URL

### 分析

- **Deck / Table** on `v2_Stats`（待審核、銷售額、已入場…）
- **Pie Chart** on `v2_Chart`  
  - Chart columns **只勾** `label` + `revenue`（或 `sales_qty`）  
  - **不要勾** `_RowNumber`

### 郵件

- 按鈕「寄出郵件」→ `as_action=寄出郵件` → `processV2Actions`  
- 使用 `GmailApp`（試算表擁有者 Gmail）  
- 可在 Settings 設 `EMAIL_REPLY_TO`、`V2_EMAIL_FROM_NAME`

---

## 四、訂單狀態機

```
Tally 新單
   ↓ sync
待審核 ──拒絕──► 已拒絕
   │
   └─通過──► 已通過 ──(票務自動)──► 已出票
                │
                └─寄出郵件──► receipt_status=已寄出

v2_Tickets.status:
  有效 ──掃碼──► 已入場
  可 作廢
```

---

## 五、與舊系統關係

| 舊 | 2.0 |
|----|-----|
| NA_Tickets / NA_Merch | 仍是 Tally 原始來源（不刪） |
| Order_Categories | 品項／hint／單價主檔 |
| 挪亞審核 / AS_Pie_* | 可並存；App 改連 **v2_*** |
| NOAH控制器 | 可繼續用；手機主用 AppSheet v2 |

---

## 六、驗收清單

- [ ] `installNoahArkV2` 成功，六張 v2 表有表頭與資料  
- [ ] 手機打開 App 看到待審核訂單  
- [ ] 點開訂單看得到**付款截圖**  
- [ ] 按「通過」後票務出現在 `v2_Tickets` 且有 QR 圖  
- [ ] 掃 QR 回傳 CONFIRMED，狀態變已入場  
- [ ] 「寄出郵件」顧客收到信  
- [ ] `v2_Stats` / Pie 有銷量與銷售額  

---

## 七、安全提醒

- Web App「任何人」可呼叫時，請只暴露 scan API，勿公開試算表  
- 正式場地可改：僅網域內使用者 + 密碼 Slice  
- 付款截圖 URL 若為 Tally 私有連結，可能過期；重要單建議另存 Drive  
