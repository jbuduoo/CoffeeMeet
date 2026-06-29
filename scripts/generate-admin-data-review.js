const fs = require("fs");
const { getSheetValues } = require("../sheets-client");

function rowsToObjects(values = []) {
  const [headers = [], ...rows] = values;
  return rows
    .filter((row) => row.some((cell) => String(cell || "").trim()))
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] || ""])));
}

function value(row, keys) {
  for (const key of keys) {
    const found = row[key];
    if (found !== undefined && found !== null && String(found).trim()) return String(found).trim();
  }
  return "";
}

function normalizeStatus(status) {
  const normalized = String(status || "").trim().toLowerCase();
  if (["pending", "sent", "已邀約"].includes(normalized)) return "sent";
  if (["incoming", "被邀約"].includes(normalized)) return "incoming";
  if (["confirmed", "active", "accepted", "scheduled", "ready", "待見面", "邀約成立"].includes(normalized)) return "confirmed";
  if (["done", "completed", "complete", "finished", "已完成"].includes(normalized)) return "done";
  if (["cancelled", "canceled", "declined", "rejected", "已取消"].includes(normalized)) return "cancelled";
  return normalized || "sent";
}

function toNumber(input) {
  const number = Number(input);
  return Number.isFinite(number) ? number : null;
}

function distanceKm(a, b) {
  if (!a?.lat || !a?.lng || !b?.lat || !b?.lng) return null;
  const lat1 = toNumber(a.lat);
  const lng1 = toNumber(a.lng);
  const lat2 = toNumber(b.lat);
  const lng2 = toNumber(b.lng);
  if ([lat1, lng1, lat2, lng2].some((item) => item === null)) return null;
  const toRad = (degree) => (degree * Math.PI) / 180;
  const radius = 6371;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const h =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return Math.round(2 * radius * Math.asin(Math.sqrt(h)) * 10) / 10;
}

function taipeiDateKey(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Taipei",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(date);
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function displayName(user) {
  if (!user) return "未記錄";
  return user.name || "未填姓名";
}

function displayUser(user) {
  if (!user) return "未記錄";
  const email = user.email ? ` / ${user.email}` : "";
  return `${displayName(user)}${email}`;
}

function adminTestRecipient() {
  return {
    name: "管理者測試",
    email: "jbuduoo@gmail.com",
    userId: "admin-test-jbuduoo",
    missing: ["測試寄送"],
    isAdminTest: true,
  };
}

function formatMissing(fields) {
  return fields.filter((field) => !field.value).map((field) => field.label);
}

function hasPublicMeetupKeyword(placeName) {
  const normalized = String(placeName || "").toLowerCase();
  return [
    "全家",
    "7-11",
    "7－11",
    "711",
    "seven",
    "星巴克",
    "starbucks",
    "咖啡",
    "coffee",
    "cafe",
    "café",
  ].some((keyword) => normalized.includes(keyword));
}

function locationStatus(user, place) {
  if (!place && !user.meeting_place_id) return "缺見面地點";
  if (!place) return "找不到 meeting_place_id";
  if (!place.lat || !place.lng) return "有地點但缺座標";
  if (!place.map_url) return "有座標但缺 Google Maps URL";
  if (!hasPublicMeetupKeyword(place.place_name)) return "疑似不是公開店點";
  return "有座標與 Google Maps URL";
}

function escapeScriptJson(data) {
  return JSON.stringify(data, null, 2).replace(/</g, "\\u003c");
}

function escapeTemplateScript(value) {
  return String(value).replace(/\\/g, "\\\\").replace(/`/g, "\\`").replace(/\$\{/g, "\\${");
}

function buildHtml(data) {
  return `<!doctype html>
<html lang="zh-Hant">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>CoffeeMeet 管理資料盤點</title>
    <style>
      :root {
        --bg: #f7f5f0;
        --panel: #ffffff;
        --soft: #fbfaf7;
        --text: #28231f;
        --muted: #716b63;
        --line: #ddd6ca;
        --good: #0f766e;
        --warn: #b45309;
        --bad: #b91c1c;
      }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        background: var(--bg);
        color: var(--text);
        font-family: Inter, "Noto Sans TC", system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      }
      main {
        width: min(1280px, calc(100% - 32px));
        margin: 0 auto;
        padding: 28px 0 40px;
      }
      header {
        display: flex;
        justify-content: space-between;
        align-items: end;
        gap: 16px;
        margin-bottom: 14px;
      }
      h1 { margin: 0; font-size: 28px; letter-spacing: 0; }
      .subhead { color: var(--muted); font-size: 14px; margin: 6px 0 0; }
      .tabs, .filters {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-bottom: 12px;
      }
      button, input, select {
        min-height: 40px;
        border: 1px solid var(--line);
        border-radius: 8px;
        background: var(--panel);
        color: var(--text);
        font: inherit;
      }
      button {
        cursor: pointer;
        padding: 0 14px;
        font-weight: 800;
      }
      button.active {
        border-color: var(--good);
        background: #e6f4f1;
        color: #115e59;
      }
      input, select { padding: 0 12px; }
      input { min-width: 260px; }
      .summary {
        display: grid;
        grid-template-columns: repeat(6, minmax(0, 1fr));
        gap: 8px;
        margin-bottom: 12px;
      }
      .stat, .panel {
        border: 1px solid var(--line);
        border-radius: 8px;
        background: var(--panel);
      }
      .stat { padding: 12px; }
      .stat span { display: block; color: var(--muted); font-size: 12px; margin-bottom: 4px; }
      .stat strong { display: block; font-size: 24px; }
      .panel { overflow: hidden; }
      .admin-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 12px;
      }
      .admin-block {
        border: 1px solid var(--line);
        border-radius: 8px;
        background: var(--panel);
        overflow: hidden;
      }
      .admin-block h2 {
        margin: 0;
        padding: 12px 14px;
        border-bottom: 1px solid var(--line);
        background: var(--soft);
        font-size: 16px;
      }
      .admin-block table {
        min-width: 0;
      }
      .table-wrap { overflow-x: auto; }
      table { width: 100%; min-width: 1080px; border-collapse: collapse; }
      th, td {
        padding: 10px 12px;
        border-bottom: 1px solid var(--line);
        text-align: left;
        vertical-align: top;
        font-size: 13px;
      }
      th {
        position: sticky;
        top: 0;
        z-index: 1;
        background: var(--soft);
        color: var(--muted);
        font-weight: 900;
      }
      tr:last-child td { border-bottom: 0; }
      tbody tr:hover { background: #fffdf8; }
      .pill {
        display: inline-flex;
        align-items: center;
        min-height: 24px;
        border-radius: 999px;
        padding: 0 9px;
        background: #e6f4f1;
        color: #115e59;
        font-size: 12px;
        font-weight: 900;
      }
      .pill.warn { background: #fef3c7; color: var(--warn); }
      .pill.bad { background: #fee2e2; color: var(--bad); }
      .missing { color: var(--bad); font-weight: 800; }
      .mail-body {
        white-space: pre-line;
        line-height: 1.75;
      }
      .send-reminder {
        border: 1px solid #f4c7c3;
        border-radius: 8px;
        background: #fff5f5;
        color: var(--bad);
        font-weight: 800;
        line-height: 1.6;
        margin-bottom: 12px;
        padding: 12px 14px;
      }
      .send-checkbox {
        width: 14px;
        height: 14px;
        min-height: 14px;
        margin: 0;
        accent-color: var(--bad);
      }
      .check-cell {
        text-align: center;
        width: 48px;
      }
      .status-cell {
        width: 72px;
      }
      .red {
        color: var(--bad);
        font-weight: 900;
      }
      .view.hidden { display: none; }
      @media (max-width: 860px) {
        header { align-items: stretch; flex-direction: column; }
        .summary { grid-template-columns: repeat(2, minmax(0, 1fr)); }
        .admin-grid { grid-template-columns: 1fr; }
        input, select, button { width: 100%; }
      }
    </style>
  </head>
  <body>
    <main>
      <header>
        <div>
          <h1>CoffeeMeet 管理資料盤點</h1>
          <p class="subhead">產生時間：${data.generatedAt}。這份資料直接寫在 HTML 裡，方便先人工檢查與修改。</p>
        </div>
      </header>
      <nav class="tabs">
        <button class="active" type="button" data-view="overview">總覽</button>
        <button type="button" data-view="users">使用者資料</button>
        <button type="button" data-view="missing">缺資料名單</button>
        <button type="button" data-view="mailSamples">補資料 Email 樣本</button>
        <button type="button" data-view="invites">邀約與距離</button>
        <button type="button" data-view="locations">定位檢查</button>
      </nav>
      <section class="filters">
        <input id="searchInput" type="search" placeholder="搜尋姓名、Email、User ID、邀約 ID、地點" />
        <select id="genderFilter">
          <option value="all">全部性別</option>
          <option value="男性">男性</option>
          <option value="女性">女性</option>
          <option value="未填">未填性別</option>
        </select>
      </section>
      <section class="summary" id="summary"></section>
      <section class="view" id="overviewView"><div class="admin-grid" id="adminBlocks"></div></section>
      <section class="view hidden" id="usersView"><div class="panel"><div class="table-wrap"><table><thead><tr><th>姓名</th><th>Email</th><th>性別</th><th>年齡</th><th>地區</th><th>職業</th><th>學歷</th><th>可約時間</th><th>見面地點</th><th>缺資料</th><th>User ID</th></tr></thead><tbody id="userRows"></tbody></table></div></div></section>
      <section class="view hidden" id="missingView"><div class="send-reminder">寄送提醒：請先確認名單與缺資料項目，只勾選確定要寄送的人。之後接寄信功能時，系統只會寄給已勾選的人。</div><div class="panel"><div class="table-wrap"><table><thead><tr><th class="check-cell">寄送</th><th>姓名</th><th>Email</th><th>性別</th><th>缺資料</th><th>建議處理</th><th class="status-cell">狀態</th><th>User ID</th></tr></thead><tbody id="missingRows"></tbody></table></div></div></section>
      <section class="view hidden" id="mailSamplesView"><div class="admin-grid" id="mailSampleBlocks"></div></section>
      <section class="view hidden" id="invitesView"><div class="panel"><div class="table-wrap"><table><thead><tr><th>誰邀請</th><th>邀請誰</th><th>距離</th><th>狀態</th><th>是否完成</th><th>時間</th><th>地點</th><th>邀約 ID</th></tr></thead><tbody id="inviteRows"></tbody></table></div></div></section>
      <section class="view hidden" id="locationsView"><div class="panel"><div class="table-wrap"><table><thead><tr><th>姓名</th><th>Email</th><th>見面地點</th><th>定位狀態</th><th>座標</th><th>Google Maps</th><th>User ID</th></tr></thead><tbody id="locationRows"></tbody></table></div></div></section>
    </main>
    <script>
      const ADMIN_DATA = ${escapeScriptJson(data)};
      const state = { view: "overview", query: "", gender: "all" };
      const statusLabels = { sent: "已邀約", incoming: "被邀約", confirmed: "待見面", done: "已完成", cancelled: "已取消" };
      function esc(value) {
        return String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
      }
      function matchesSearch(item) {
        const q = state.query;
        if (!q) return true;
        return Object.values(item).join(" ").toLowerCase().includes(q);
      }
      function matchesGender(user) {
        if (state.gender === "all") return true;
        if (state.gender === "未填") return !user.gender;
        return user.gender === state.gender;
      }
      function filteredUsers() {
        return ADMIN_DATA.users.filter((user) => matchesGender(user) && matchesSearch(user));
      }
      function filteredInvites() {
        return ADMIN_DATA.invites.filter((invite) => matchesSearch(invite));
      }
      function renderSummary() {
        const users = ADMIN_DATA.users;
        const invites = ADMIN_DATA.invites;
        const stats = [
          ["男生", users.filter((u) => u.gender === "男性").length],
          ["女生", users.filter((u) => u.gender === "女性").length],
          ["缺資料", users.filter((u) => u.missing.length).length],
          ["缺定位", users.filter((u) => !u.locationOk).length],
          ["邀約數", invites.length],
          ["已完成", invites.filter((i) => i.completed).length],
        ];
        document.getElementById("summary").innerHTML = stats.map(([label, value]) => '<div class="stat"><span>' + esc(label) + '</span><strong>' + esc(value) + '</strong></div>').join("");
      }
      function renderOverview() {
        const sectionDefs = [
          ["今日新增人數", "看宣傳有沒有效", ADMIN_DATA.sections.todayNewUsers, ["姓名", "性別/地區", "缺資料"]],
          ["今日新增邀約", "看互動有沒有發生", ADMIN_DATA.sections.todayNewInvites, ["誰邀請誰", "狀態", "距離"]],
          ["待回覆邀約", "看誰卡住", ADMIN_DATA.sections.pendingInvites, ["誰邀請誰", "建立時間", "距離"]],
          ["明天待見面", "方便提醒", ADMIN_DATA.sections.tomorrowMeetups, ["誰邀請誰", "時間", "地點"]],
          ["缺資料名單", "方便催補", ADMIN_DATA.sections.missingUsers, ["姓名/Email", "缺資料", "定位"]],
          ["高活躍使用者", "可以人工關心", ADMIN_DATA.sections.activeUsers, ["使用者", "邀約數", "完成數"]],
          ["近距離可配對名單", "管理員可協助撮合", ADMIN_DATA.sections.nearbyPairs, ["男生", "女生", "距離"]],
        ];
        document.getElementById("adminBlocks").innerHTML = sectionDefs.map(([title, purpose, rows, headers]) => {
          const body = rows.length
            ? rows.slice(0, 8).map((row) => '<tr>' + row.cells.map((cell) => '<td>' + esc(cell) + '</td>').join('') + '</tr>').join('')
            : '<tr><td colspan="' + headers.length + '">目前沒有資料</td></tr>';
          return '<article class="admin-block"><h2>' + esc(title) + ' <span class="subhead">(' + esc(purpose) + ')</span></h2><div class="table-wrap"><table><thead><tr>' + headers.map((header) => '<th>' + esc(header) + '</th>').join('') + '</tr></thead><tbody>' + body + '</tbody></table></div></article>';
        }).join("");
      }
      function missingHtml(missing) {
        return missing.length ? '<span class="missing">' + esc(missing.join("、")) + '</span>' : '<span class="pill">完整</span>';
      }
      function renderUsers() {
        document.getElementById("userRows").innerHTML = filteredUsers().map((user) => "<tr><td>" + esc(user.name || "未填") + "</td><td>" + esc(user.email) + "</td><td>" + esc(user.gender || "未填") + "</td><td>" + esc(user.age || "未填") + "</td><td>" + esc(user.area || "未填") + "</td><td>" + esc(user.occupation || "未填") + "</td><td>" + esc(user.education || "未填") + "</td><td>" + esc(user.availableTimes || "未填") + "</td><td>" + esc(user.placeName || "未填") + "</td><td>" + missingHtml(user.missing) + "</td><td>" + esc(user.userId) + "</td></tr>").join("");
      }
      function renderMissing() {
        const rows = filteredUsers().filter((user) => user.missing.length);
        document.getElementById("missingRows").innerHTML = rows.map((user) => \`<tr><td class="check-cell"><input class="send-checkbox" type="checkbox" data-send-user="\${esc(user.userId)}" aria-label="勾選寄送給 \${esc(user.email || user.userId)}" /></td><td>\${esc(user.name || "未填")}</td><td>\${esc(user.email)}</td><td>\${esc(user.gender || "未填")}</td><td>\${missingHtml(user.missing)}</td><td>可寄信請他補：\${esc(user.missing.join("、"))}</td><td class="status-cell"><span class="pill warn">未寄</span></td><td>\${esc(user.userId)}</td></tr>\`).join("");
      }
      function renderMailSamples() {
        document.getElementById("mailSampleBlocks").innerHTML = ADMIN_DATA.mailSamples.map((sample) => {
          const users = sample.users.length
            ? sample.users.slice(0, 10).map((user) => '<tr><td class="check-cell"><input class="send-checkbox" type="checkbox" data-send-user="' + esc(user.userId) + '" aria-label="勾選寄送給 ' + esc(user.email || user.userId) + '" /></td><td>' + esc(user.name || "未填") + '</td><td>' + esc(user.email || "未填") + '</td><td>' + esc(user.missing.join("、")) + '</td><td class="status-cell"><span class="pill warn">未寄</span></td></tr>').join("")
            : '<tr><td colspan="5">目前沒有這類名單</td></tr>';
          const note = sample.note ? '<tr><th>使用方式</th><td colspan="2">' + esc(sample.note) + '</td></tr>' : '';
          const bodyHtml = sample.bodyHtml || esc(sample.body);
          return '<article class="admin-block"><h2>' + esc(sample.title) + ' <span class="subhead">(' + sample.users.length + ' 人)</span></h2><div class="send-reminder">寄送提醒：請先確認文案與名單，只勾選確定要寄送的人。現在只是樣本，不會真的寄出。</div><div class="table-wrap"><table><tbody>' + note + '<tr><th>建議主旨</th><td colspan="2">' + esc(sample.subject) + '</td></tr><tr><th>信件內容</th><td colspan="2"><div class="mail-body">' + bodyHtml + '</div></td></tr></tbody></table><table><thead><tr><th class="check-cell">寄送</th><th>姓名</th><th>Email</th><th>缺資料</th><th class="status-cell">狀態</th></tr></thead><tbody>' + users + '</tbody></table></div></article>';
        }).join("");
      }
      function renderInvites() {
        document.getElementById("inviteRows").innerHTML = filteredInvites().map((invite) => '<tr><td>' + esc(invite.senderName) + '</td><td>' + esc(invite.receiverName) + '</td><td>' + esc(invite.distanceKmText) + '</td><td><span class="pill">' + esc(statusLabels[invite.status] || invite.status) + '</span></td><td>' + (invite.completed ? '<span class="pill">完成</span>' : '<span class="pill warn">未完成</span>') + '</td><td>' + esc(invite.time || "未記錄") + '</td><td>' + esc(invite.place || "未記錄") + '</td><td>' + esc(invite.inviteId) + '</td></tr>').join("");
      }
      function renderLocations() {
        document.getElementById("locationRows").innerHTML = filteredUsers().map((user) => '<tr><td>' + esc(user.name) + '</td><td>' + esc(user.email) + '</td><td>' + esc(user.placeName || "未填") + '</td><td><span class="pill ' + (user.locationOk ? '' : 'bad') + '">' + esc(user.locationStatus) + '</span></td><td>' + esc(user.lat && user.lng ? user.lat + ", " + user.lng : "缺座標") + '</td><td>' + (user.mapUrl ? '<a href="' + esc(user.mapUrl) + '" target="_blank" rel="noreferrer">Google Maps</a>' : '缺連結') + '</td><td>' + esc(user.userId) + '</td></tr>').join("");
      }
      function render() {
        renderSummary();
        renderOverview();
        renderUsers();
        renderMissing();
        renderMailSamples();
        renderInvites();
        renderLocations();
      }
      document.querySelectorAll("[data-view]").forEach((button) => {
        button.addEventListener("click", () => {
          state.view = button.dataset.view;
          document.querySelectorAll("[data-view]").forEach((item) => item.classList.toggle("active", item === button));
          document.querySelectorAll(".view").forEach((view) => view.classList.add("hidden"));
          document.getElementById(state.view + "View").classList.remove("hidden");
        });
      });
      document.getElementById("searchInput").addEventListener("input", (event) => {
        state.query = event.target.value.trim().toLowerCase();
        render();
      });
      document.getElementById("genderFilter").addEventListener("change", (event) => {
        state.gender = event.target.value;
        render();
      });
      render();
    </script>
  </body>
</html>
`;
}

async function main() {
  const [usersSheet, placesSheet, photosSheet, invitesSheet] = await Promise.all([
    getSheetValues("users!A1:AZ1000"),
    getSheetValues("meeting_places!A1:AZ1000"),
    getSheetValues("user_photos!A1:AZ1000"),
    getSheetValues("invites!A1:AZ1000"),
  ]);

  const rawUsers = rowsToObjects(usersSheet.values);
  const rawPlaces = rowsToObjects(placesSheet.values);
  const rawPhotos = rowsToObjects(photosSheet.values);
  const rawInvites = rowsToObjects(invitesSheet.values);
  const placesById = new Map(rawPlaces.map((place) => [place.place_id, place]));
  const placesByUserId = new Map(rawPlaces.filter((place) => place.owner_user_id).map((place) => [place.owner_user_id, place]));
  const photosByUserId = rawPhotos.reduce((groups, photo) => {
    if (!photo.user_id || String(photo.status || "") === "deleted") return groups;
    groups[photo.user_id] ||= [];
    groups[photo.user_id].push(photo);
    return groups;
  }, {});

  const users = rawUsers
    .filter((user) => user.status !== "inactive")
    .map((user) => {
      const place = placesById.get(user.meeting_place_id) || placesByUserId.get(user.user_id) || null;
      const isPublicMeetupPlace = hasPublicMeetupKeyword(place?.place_name || "");
      const missing = formatMissing([
        { label: "Email", value: user.email },
        { label: "姓名", value: user.nickname },
        { label: "性別", value: user.gender },
        { label: "年齡", value: user.age || user.birth_year },
        { label: "地區", value: user.city && user.district },
        { label: "職業", value: user.occupation },
        { label: "學歷", value: user.education },
        { label: "可約時間", value: user.available_times },
        { label: "見面地點", value: place?.place_name || user.meeting_place_id },
        { label: "Google Maps定位", value: place?.lat && place?.lng && place?.map_url },
        { label: "公開店點定位", value: !place || !place.place_name ? true : isPublicMeetupPlace },
        { label: "自我介紹", value: user.intro },
        { label: "照片", value: photosByUserId[user.user_id]?.length },
      ]);
      const locStatus = locationStatus(user, place);
      return {
        userId: user.user_id || "",
        name: user.nickname || "",
        email: user.email || "",
        gender: user.gender || "",
        age: user.age || "",
        area: [user.city, user.district].filter(Boolean).join("・"),
        occupation: user.occupation || "",
        education: user.education || "",
        relationshipStatus: user.relationship_status || "",
        availableTimes: user.available_times || "",
        placeName: place?.place_name || user.meeting_place_id || "",
        mapUrl: place?.map_url || "",
        lat: place?.lat || "",
        lng: place?.lng || "",
        intro: user.intro || "",
        photoCount: photosByUserId[user.user_id]?.length || 0,
        createdAt: user.created_at || user.updated_at || "",
        missing,
        locationStatus: locStatus,
        locationOk: locStatus === "有座標與 Google Maps URL",
      };
    });

  const usersById = new Map(users.map((user) => [user.userId, user]));
  const invites = rawInvites.map((invite, index) => {
    const senderId = value(invite, ["sender_user_id", "senderUserId", "sender_id"]);
    const receiverId = value(invite, ["receiver_user_id", "receiverUserId", "receiver_id"]);
    const sender = usersById.get(senderId);
    const receiver = usersById.get(receiverId);
    const distance = distanceKm(sender, receiver);
    const status = normalizeStatus(invite.status);
    return {
      inviteId: invite.invite_id || `invite-row-${index + 1}`,
      senderId,
      receiverId,
      senderName: sender ? `${sender.name} / ${sender.email}` : senderId || "未記錄",
      receiverName: receiver ? `${receiver.name} / ${receiver.email}` : receiverId || "未記錄",
      status,
      completed: status === "done" || Boolean(invite.feedback_score || invite.male_feedback_score || invite.female_feedback_score),
      distanceKm: distance,
      distanceKmText: distance === null ? "無法計算" : `${distance} km`,
      time: invite.accepted_at || invite.selected_times || "",
      place: invite.place_name || receiver?.placeName || sender?.placeName || "",
      createdAt: invite.created_at || "",
      updatedAt: invite.updated_at || "",
    };
  });

  const distances = invites.map((invite) => invite.distanceKm).filter((item) => item !== null);
  const now = new Date();
  const todayKey = taipeiDateKey(now);
  const tomorrowKey = taipeiDateKey(addDays(now, 1));
  const todayNewUsers = users
    .filter((user) => taipeiDateKey(user.createdAt) === todayKey)
    .map((user) => ({
      cells: [
        displayName(user),
        [user.gender || "未填性別", user.area || "未填地區"].join(" / "),
        user.missing.length ? user.missing.join("、") : "完整",
      ],
    }));
  const todayNewInvites = invites
    .filter((invite) => taipeiDateKey(invite.createdAt) === todayKey)
    .map((invite) => ({
      cells: [
        `${displayUser(usersById.get(invite.senderId))} → ${displayUser(usersById.get(invite.receiverId))}`,
        normalizeStatus(invite.status),
        invite.distanceKmText,
      ],
    }));
  const pendingInvites = invites
    .filter((invite) => invite.status === "sent" || invite.status === "incoming")
    .map((invite) => ({
      cells: [
        `${displayUser(usersById.get(invite.senderId))} → ${displayUser(usersById.get(invite.receiverId))}`,
        invite.createdAt || "未記錄",
        invite.distanceKmText,
      ],
    }));
  const tomorrowMeetups = invites
    .filter((invite) => invite.status === "confirmed" && taipeiDateKey(invite.time) === tomorrowKey)
    .map((invite) => ({
      cells: [
        `${displayUser(usersById.get(invite.senderId))} → ${displayUser(usersById.get(invite.receiverId))}`,
        invite.time || "未記錄",
        invite.place || "未記錄",
      ],
    }));
  const missingUsers = users
    .filter((user) => user.missing.length)
    .sort((a, b) => b.missing.length - a.missing.length)
    .map((user) => ({
      cells: [
        displayUser(user),
        user.missing.join("、"),
        user.locationStatus,
      ],
    }));
  const activeUsers = users
    .map((user) => {
      const related = invites.filter((invite) => invite.senderId === user.userId || invite.receiverId === user.userId);
      return {
        user,
        total: related.length,
        sent: related.filter((invite) => invite.senderId === user.userId).length,
        received: related.filter((invite) => invite.receiverId === user.userId).length,
        completed: related.filter((invite) => invite.completed).length,
      };
    })
    .filter((item) => item.total > 0)
    .sort((a, b) => b.total - a.total || b.completed - a.completed)
    .map((item) => ({
      cells: [
        displayUser(item.user),
        `${item.total} 筆（發出 ${item.sent} / 收到 ${item.received}）`,
        `${item.completed} 筆`,
      ],
    }));
  const existingPairKeys = new Set(invites.map((invite) => [invite.senderId, invite.receiverId].sort().join("|")));
  const males = users.filter((user) => user.gender === "男性" && user.locationOk);
  const females = users.filter((user) => user.gender === "女性" && user.locationOk);
  const nearbyPairs = males
    .flatMap((male) => females.map((female) => {
      const distance = distanceKm(male, female);
      return { male, female, distance };
    }))
    .filter((pair) => pair.distance !== null && pair.distance <= 10)
    .filter((pair) => !existingPairKeys.has([pair.male.userId, pair.female.userId].sort().join("|")))
    .sort((a, b) => a.distance - b.distance)
    .slice(0, 30)
    .map((pair) => ({
      cells: [
        displayUser(pair.male),
        displayUser(pair.female),
        `${pair.distance} km`,
      ],
    }));
  const mailSamples = [
    {
      title: "基本版：請補齊資料",
      note: "目前建議先用這封。適合所有缺資料的人，不用先分太細，也避免同一個人收到多封提醒。",
      subject: "請補齊 CoffeeMeet 資料，讓別人更容易邀你喝咖啡",
      body: "Hi {{name}}，\n\n你的 CoffeeMeet 名片已經建立了，但目前還有一些資料可以補完整：\n\n{{missingList}}\n\n資料越完整，別人越容易判斷你是真人，也更知道要不要邀你喝咖啡。\n\n特別提醒兩件事：\n\n1. 照片很重要\n有照片的人比較容易被邀約，也比較像真人。建議先放 1 張清楚生活照，不需要修圖，只要自然、看得清楚就可以。\n\n2. 見面地點要明確，並通過定位驗證\n請依照頁面圖片上的指示操作：搜尋一間公開店點，選 Google Maps 上找得到的店家，把完整店名貼上，並通過驗證。\n\n建議選全家、7-11、星巴克，或 Google Maps 上找得到的咖啡店。\n\n因為這是男生來拜訪你的地點；如果你不安心，可以稍微向外 100 公尺，找個公開的便利商店。\n\n如果你是男生，這個地點會用來計算你和女生之間的距離。\n\n不要只填行政區、路名、社區或不明確的位置，對方會不知道實際要去哪裡，也比較難放心邀約。\n\n不用一次寫很多，先把照片、自我介紹、可約時間、公開見面地點補上，就可以讓邀約成功率提高很多。\n\n你也已被邀請加入「一杯咖啡的時間，認識一個人(交友)」社群。\n請點選以下連結加入：\nhttps://line.me/ti/g2/Ac5c0VtixLIbZmRur9NsI0A3P6_CgsJmUEUcnw?utm_source=invitation&utm_medium=link_copy&utm_campaign=default",
      bodyHtml: "Hi {{name}}，\n\n你的 CoffeeMeet 名片已經建立了，但目前還有一些資料可以補完整：\n\n<span class=\"red\">{{missingList}}</span>\n\n資料越完整，別人越容易判斷你是真人，也更知道要不要邀你喝咖啡。\n\n特別提醒兩件事：\n\n<span class=\"red\">1. 照片很重要</span>\n有照片的人比較容易被邀約，也比較像真人。建議先放 1 張清楚生活照，不需要修圖，只要自然、看得清楚就可以。\n\n<span class=\"red\">2. 見面地點要明確，並通過定位驗證</span>\n請依照頁面圖片上的指示操作：搜尋一間公開店點，選 Google Maps 上找得到的店家，把完整店名貼上，並通過驗證。\n\n<span class=\"red\">建議選全家、7-11、星巴克，或 Google Maps 上找得到的咖啡店。</span>\n\n因為這是男生來拜訪你的地點；如果你不安心，可以稍微向外 100 公尺，找個公開的便利商店。\n\n如果你是男生，這個地點會用來計算你和女生之間的距離。\n\n<span class=\"red\">不要只填行政區、路名、社區或不明確的位置。</span> 對方會不知道實際要去哪裡，也比較難放心邀約。\n\n不用一次寫很多，先把照片、自我介紹、可約時間、公開見面地點補上，就可以讓邀約成功率提高很多。\n\n你也已被邀請加入「一杯咖啡的時間，認識一個人(交友)」社群。\n請點選以下連結加入：\n<a href=\"https://line.me/ti/g2/Ac5c0VtixLIbZmRur9NsI0A3P6_CgsJmUEUcnw?utm_source=invitation&utm_medium=link_copy&utm_campaign=default\" target=\"_blank\" rel=\"noreferrer\">https://line.me/ti/g2/Ac5c0VtixLIbZmRur9NsI0A3P6_CgsJmUEUcnw?utm_source=invitation&utm_medium=link_copy&utm_campaign=default</a>",
      users: [adminTestRecipient(), ...users.filter((user) => user.missing.length)],
    },
  ].map((sample) => ({
    ...sample,
    users: sample.users.map((user) => ({
      name: user.name,
      email: user.email,
      userId: user.userId,
      missing: user.missing,
    })),
  }));
  const data = {
    generatedAt: new Date().toLocaleString("zh-TW", { timeZone: "Asia/Taipei" }),
    users,
    invites,
    sections: {
      todayNewUsers,
      todayNewInvites,
      pendingInvites,
      tomorrowMeetups,
      missingUsers,
      activeUsers,
      nearbyPairs,
    },
    mailSamples,
    rawCounts: {
      users: rawUsers.length,
      places: rawPlaces.length,
      photos: rawPhotos.length,
      invites: rawInvites.length,
    },
    averageInviteDistanceKm: distances.length
      ? `${Math.round((distances.reduce((sum, item) => sum + item, 0) / distances.length) * 10) / 10} km`
      : "",
  };

  fs.writeFileSync("admin-data-review.html", buildHtml(data), "utf8");
  console.log(`Wrote admin-data-review.html with ${users.length} users and ${invites.length} invites.`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
