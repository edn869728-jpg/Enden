# Enden 免費叫車系統

本系統為免費叫車紀錄平台，透過 Google Apps Script 驅動，無需安裝任何 App，直接在網頁上填寫叫車需求、查詢歷史叫車紀錄，輕鬆管理每一趟行程。

## 🚗 立即使用

點擊下方連結即可開啟叫車系統：

**[➜ 開啟叫車系統](https://script.google.com/macros/s/AKfycby19SkYJuAtue75pJ0nPSR_HEudfNV5t17F9eZwmmdNjcDM2aEYrV_xKHVyQx2ToLSWOA/exec)**

## 功能說明

- 📝 **填寫叫車**：輸入起點、終點及乘車時間，快速完成叫車申請
- 🔍 **查詢紀錄**：隨時查詢個人或全部叫車歷史紀錄
- 📍 **多點行程**：支援多個中途停靠點，彈性安排行程路線
- 💬 **LINE 整合**：可搭配 LINE Bot 進行叫車互動與即時通知

## 技術架構

- Google Apps Script Web App
- Google 試算表（資料儲存）
- LINE Messaging API（Bot 互動）
- 純 HTML / CSS / JavaScript 前端

## 其他入口

| 頁面 | 說明 |
|------|------|
| [index.html](./index.html) | 系統說明與叫車入口頁 |
| [employee.html](./employee.html) | 員工自助系統 |
| [admin.html](./admin.html) | 管理後台 |
