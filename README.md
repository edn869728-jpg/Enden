# LINE Ride - Apps Script 強化完整覆蓋版 v1.1

LINE Ride 是一個以 Google Apps Script 為後端的叫車服務，使用 Google Sheets 作為資料庫，並整合 LINE Messaging API。

## 功能

- **訂單管理**：建立、取消、查詢叫車訂單
- **常用地點**：儲存、刪除、列出常用上下車地點
- **黑名單**：新增與列出黑名單使用者
- **LINE Webhook**：處理 follow / unfollow / 文字訊息 / postback 事件
- **設定管理**：透過 Settings 工作表設定系統參數

## 部署方式

1. 在 Google Apps Script 專案中貼上 `Code.gs` 內容
2. 設定 `SHEET_ID` 為你的 Google Sheets 試算表 ID
3. 執行 `testInitSystem()` 初始化所有工作表與預設設定
4. 在 Settings 工作表填入：
   - `line_channel_access_token`
   - `liff_id`
   - `app_url`
5. 發布為 Web 應用程式，並將 Webhook URL 設定至 LINE Developers Console

## API Actions

| Action | 方法 | 說明 |
|---|---|---|
| `health` | GET | 健康檢查 |
| `init` | GET/POST | 初始化系統 |
| `config` | GET/POST | 取得公開設定 |
| `ping` | POST | Ping 測試 |
| `create_order` | POST | 建立訂單 |
| `cancel_order` | POST | 取消訂單 |
| `list_orders` | POST | 查詢訂單列表 |
| `save_favorite` | POST | 儲存常用地點 |
| `delete_favorite` | POST | 刪除常用地點 |
| `list_favorites` | POST | 查詢常用地點 |
| `blacklist_add` | POST | 新增黑名單 |
| `blacklist_list` | POST | 查詢黑名單 |
