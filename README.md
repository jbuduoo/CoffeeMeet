# CoffeeMeet

一杯咖啡的時間，用一小時認識一個人。

## 目前架構

- GitHub Pages：放前端頁面 `index.html`
- Google Apps Script：處理資料讀寫
- Google Sheets：存會員、邀約、地點、照片資料
- Google Drive：存照片

Render / Node 版本先保留在 repo 裡，但正式展示可以優先用 GitHub Pages，避免 Render 免費服務休眠造成使用者打不開。

## GitHub Pages 網址

https://jbuduoo.github.io/CoffeeMeet/

## Google 資料

- Google Sheet：`1JzQwlNWQphrHfUCcpQl3pC6jHLXN6ykwvJKSs5Uaxsg`
- BOY 照片資料夾：`12M91XLYJ7seBNhrsGL2JZNJZlSGlL9ri`
- GIRL 照片資料夾：`1r_6bPvf-R5lHM92g379al71G8zzmj-of`

## Apps Script

範本在：

```text
docs/google-apps-script-backend.js
```

目前 Web App URL：

```text
https://script.google.com/macros/s/AKfycbzMQ43-vk2l4ngrdut0V9jYncaVIgTGXlqVV-9KvIiwELaXBwwPV1f7XHbtt4R_r97M5g/exec
```

部署成 Web App 後，把部署網址填到 `index.html`：

```js
const GOOGLE_APPS_SCRIPT_URL = "你的 Apps Script Web App URL";
```

如果這個值是空字串，網站會使用內建示範資料，不會呼叫後端。

## 本機預覽

直接開 `index.html` 即可看畫面。

若要使用舊的 Node / Render API 測試：

```powershell
npm install
npm start
```

再開：

```text
http://localhost:3000/
```
