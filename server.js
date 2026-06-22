const http = require("http");
const fs = require("fs");
const path = require("path");
const { getSpreadsheetMeta, getSheetValues } = require("./sheets-client");

const port = Number(process.env.PORT || 3000);
const rootDir = __dirname;

function rowsToObjects(values = []) {
  const [headers = [], ...rows] = values;
  return rows
    .filter((row) => row.some((cell) => String(cell || "").trim()))
    .map((row) =>
      Object.fromEntries(headers.map((header, index) => [header, row[index] || ""])),
    );
}

function statusToInviteStatus(status) {
  const normalized = String(status || "").toLowerCase();
  if (["incoming", "sent", "confirmed", "done", "cancelled"].includes(normalized)) return normalized;
  if (["accepted", "scheduled", "ready", "待見面"].includes(normalized)) return "confirmed";
  if (["completed", "complete", "已完成"].includes(normalized)) return "done";
  if (["canceled", "cancelled", "已取消"].includes(normalized)) return "cancelled";
  if (["pending", "sent", "已邀約"].includes(normalized)) return "sent";
  return normalized || "sent";
}

async function getAppData() {
  const [usersSheet, placesSheet, photosSheet, invitesSheet] = await Promise.all([
    getSheetValues("users!A1:Z1000"),
    getSheetValues("meeting_places!A1:Z1000"),
    getSheetValues("user_photos!A1:Z1000"),
    getSheetValues("invites!A1:Z1000"),
  ]);

  const users = rowsToObjects(usersSheet.values);
  const places = rowsToObjects(placesSheet.values);
  const photos = rowsToObjects(photosSheet.values);
  const inviteRows = rowsToObjects(invitesSheet.values);
  const placesById = Object.fromEntries(places.map((place) => [place.place_id, place]));
  const photosByUserId = photos.reduce((groups, photo) => {
    if (!photo.user_id || !photo.photo_url) return groups;
    groups[photo.user_id] ||= [];
    groups[photo.user_id].push(photo);
    return groups;
  }, {});

  const candidates = users
    .filter((user) => user.status !== "inactive")
    .map((user) => {
      const place = placesById[user.meeting_place_id] || {};
      const userPhotos = (photosByUserId[user.user_id] || []).sort(
        (a, b) => Number(a.photo_order || 99) - Number(b.photo_order || 99),
      );
      const primaryPhoto = userPhotos.find((photo) => String(photo.is_primary).toUpperCase() === "TRUE") || userPhotos[0] || {};
      return {
        id: user.user_id,
        name: user.nickname || user.user_id,
        gender: user.gender,
        age: user.age,
        area: [user.city, user.district].filter(Boolean).join("・"),
        looking: user.coffee_goal,
        occupation: user.occupation,
        education: user.education,
        relationshipStatus: user.relationship_status,
        hasChildren: user.has_children,
        smoking: user.smoking,
        drinking: user.drinking,
        intro: user.intro,
        time: user.available_times,
        place: place.place_name || user.meeting_place_id || "",
        openingQuestion: user.opening_question,
        photo: primaryPhoto.photo_url || "",
        photos: userPhotos.map((photo) => photo.photo_url),
      };
    });

  const fallbackPlace = (placeId) => placesById[placeId]?.place_name || "";
  const invites = inviteRows.map((invite, index) => ({
    id: invite.invite_id || `sheet-invite-${index + 1}`,
    candidateId: invite.receiver_user_id || invite.sender_user_id,
    status: statusToInviteStatus(invite.status),
    note: [
      invite.accepted_at ? `已確認：${invite.accepted_at}` : "",
      `地點：${fallbackPlace(invite.receiver_place_id || invite.sender_place_id)}`,
    ]
      .filter(Boolean)
      .join("；") || "從試算表載入的邀約",
  }));

  return { candidates, invites };
}

function sendJson(res, status, data) {
  res.writeHead(status, {
    "content-type": "application/json; charset=utf-8",
    "access-control-allow-origin": "*",
  });
  res.end(JSON.stringify(data, null, 2));
}

function sendFile(res, filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const contentTypes = {
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
  };
  const contentType = contentTypes[ext] || "application/octet-stream";
  res.writeHead(200, { "content-type": contentType });
  fs.createReadStream(filePath).pipe(res);
}

async function handleApi(req, res, url) {
  try {
    if (url.pathname === "/api/sheets") {
      const meta = await getSpreadsheetMeta();
      sendJson(res, 200, { ok: true, ...meta });
      return;
    }

    if (url.pathname === "/api/sheets/values") {
      const range = url.searchParams.get("range") || "A1:Z20";
      const values = await getSheetValues(range);
      sendJson(res, 200, { ok: true, range, values: values.values || [] });
      return;
    }

    if (url.pathname === "/api/app-data") {
      const data = await getAppData();
      sendJson(res, 200, { ok: true, ...data });
      return;
    }

    sendJson(res, 404, { ok: false, error: "API not found" });
  } catch (error) {
    sendJson(res, 500, { ok: false, error: error.message });
  }
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url, `http://localhost:${port}`);

  if (url.pathname.startsWith("/api/")) {
    handleApi(req, res, url);
    return;
  }

  const requestedPath = url.pathname === "/" ? "/index.html" : url.pathname;
  const safePath = path.normalize(decodeURIComponent(requestedPath)).replace(/^(\.\.[/\\])+/, "");
  const filePath = path.join(rootDir, safePath);

  if (!filePath.startsWith(rootDir) || !fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
    res.writeHead(404, { "content-type": "text/plain; charset=utf-8" });
    res.end("Not found");
    return;
  }

  sendFile(res, filePath);
});

server.listen(port, () => {
  console.log(`等一個人的咖啡 API 已啟動：http://localhost:${port}`);
  console.log(`測試試算表：http://localhost:${port}/api/sheets`);
});
