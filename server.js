const http = require("http");
const fs = require("fs");
const path = require("path");
const { appendSheetValues, getSpreadsheetMeta, getSheetValues, updateSheetValues } = require("./sheets-client");

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
  const photos = rowsToObjects(photosSheet.values).filter((photo) => photo.status !== "deleted");
  const inviteRows = rowsToObjects(invitesSheet.values);
  const placesById = Object.fromEntries(places.map((place) => [place.place_id, place]));
  const placesByOwnerUserId = Object.fromEntries(places.map((place) => [place.owner_user_id, place]).filter(([userId]) => userId));
  const photosByUserId = photos.reduce((groups, photo) => {
    if (!photo.user_id || !photo.photo_url) return groups;
    groups[photo.user_id] ||= [];
    groups[photo.user_id].push(photo);
    return groups;
  }, {});

  const candidates = users
    .filter((user) => user.status !== "inactive")
    .map((user) => {
      const place = placesById[user.meeting_place_id] || placesByOwnerUserId[user.user_id] || {};
      const userPhotos = (photosByUserId[user.user_id] || []).sort(
        (a, b) => Number(a.photo_order || 99) - Number(b.photo_order || 99),
      );
      const primaryPhoto = userPhotos.find((photo) => String(photo.is_primary).toUpperCase() === "TRUE") || userPhotos[0] || {};
      return {
        id: user.user_id,
        name: user.nickname || user.user_id,
        email: user.email,
        gender: user.gender,
        age: user.age,
        birthYear: user.birth_year || (user.age ? String(new Date().getFullYear() - Number(user.age)) : ""),
        birthMonth: user.birth_month || "",
        birthDay: user.birth_day || "",
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
        availabilityNote: user.available_times,
        place: place.place_name || user.meeting_place_id || "",
        meetingArea: place.place_name || "",
        meetingPlaceName: place.place_name || "",
        meetingPlaceUrl: place.map_url || "",
        meetingLat: place.lat || "",
        meetingLng: place.lng || "",
        openingQuestion: user.opening_question,
        interestKeywords: user.interest_keywords || user.interests || "",
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

function todayString() {
  return new Date().toISOString().slice(0, 10);
}

function calculateAge(birthYear, birthMonth, birthDay, today = new Date()) {
  const year = Number(birthYear);
  if (!year) return "";
  let age = today.getFullYear() - year;
  const month = Number(birthMonth);
  const day = Number(birthDay);
  if (month && day) {
    const birthdayThisYear = new Date(today.getFullYear(), month - 1, day);
    if (today < birthdayThisYear) age -= 1;
  }
  return String(age);
}

function slugFromEmail(email) {
  return String(email || "")
    .trim()
    .toLowerCase()
    .replace(/@.*/, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 24) || "user";
}

async function createUserFromProfile(profile) {
  const email = String(profile.email || "").trim().toLowerCase();
  if (!email) throw new Error("Email is required");

  const current = await getSheetValues("users!A1:Z1000");
  const headers = await ensureUserHeaders(current.values?.[0] || []);
  const users = rowsToObjects(current.values);
  const existingIndex = users.findIndex((user) => String(user.email || "").toLowerCase() === email);
  const existing = existingIndex >= 0 ? users[existingIndex] : null;

  const stamp = Date.now().toString(36);
  const userId = existing?.user_id || `${profile.gender === "女性" ? "girl" : "user"}-${slugFromEmail(email)}-${stamp}`;
  const placeId = existing?.meeting_place_id || `place-${slugFromEmail(email)}-${stamp}`;
  const today = todayString();
  const age = calculateAge(profile.birthYear, profile.birthMonth, profile.birthDay);

  await saveMeetingPlace({
    placeId,
    userId,
    profile,
    today,
    existingPlaceId: existing?.meeting_place_id || "",
  });

  const valuesByHeader = {
    user_id: userId,
    nickname: profile.nickname || email,
    gender: profile.gender || "",
    age,
    birth_year: profile.birthYear || "",
    birth_month: profile.birthMonth || "",
    birth_day: profile.birthDay || "",
    city: profile.city || "",
    district: profile.district || "",
    email,
    occupation: profile.occupation || "",
    education: profile.education || "",
    school: profile.school || "",
    relationship_status: profile.relationshipStatus || "",
    has_children: profile.hasChildren || "",
    coffee_goal: profile.coffeeGoal || "",
    smoking: profile.smoking || "",
    drinking: profile.drinking || "",
    intro: profile.intro || "",
    interest_keywords: profile.interestKeywords || "",
    interests: profile.interestKeywords || "",
    available_times: profile.availabilityNote || "",
    meeting_place_id: placeId,
    lock_until: existing?.lock_until || "",
    status: "active",
    created_at: existing?.created_at || today,
    updated_at: today,
    notes: existing?.notes || "",
  };

  const row = headers.map((header) => (valuesByHeader[header] !== undefined ? valuesByHeader[header] : ""));
  if (existing) {
    await updateSheetValues(`users!A${existingIndex + 2}:${columnName(headers.length)}${existingIndex + 2}`, [row]);
    const photoCount = await saveUserPhotos({ userId, profile, today });
    return { created: false, userId, photoCount };
  }

  await appendSheetValues(`users!A:${columnName(headers.length)}`, [row.length ? row : [
    userId,
    profile.nickname || email,
    profile.gender || "",
    age,
    profile.city || "",
    profile.district || "",
    email,
    profile.occupation || "",
    profile.education || "",
    profile.school || "",
    profile.relationshipStatus || "",
    profile.hasChildren || "",
    profile.coffeeGoal || "",
    profile.smoking || "",
    profile.drinking || "",
    profile.intro || "",
    profile.interestKeywords || "",
    profile.availabilityNote || "",
    "",
    placeId,
    "active",
    today,
    today,
    "",
  ]]);

  const photoCount = await saveUserPhotos({ userId, profile, today });
  return { created: true, userId, photoCount };
}

async function saveUserPhotos({ userId, profile, today }) {
  const photos = Array.isArray(profile.photos) ? profile.photos.filter(Boolean).slice(0, 3) : [];
  if (!photos.length) return 0;

  const current = await getSheetValues("user_photos!A1:Z1000");
  const headers = await ensurePhotoHeaders(current.values?.[0] || []);
  const rows = rowsToObjects(current.values);
  const existingIndexes = rows
    .map((photo, index) => ({ photo, index }))
    .filter(({ photo }) => photo.user_id === userId);

  for (let index = 0; index < photos.length; index += 1) {
    const existing = existingIndexes[index]?.photo;
    const existingRowIndex = existingIndexes[index]?.index;
    const photoUrl = photos[index];
    const valuesByHeader = {
      photo_id: existing?.photo_id || `photo-${userId}-${index + 1}`,
      user_id: userId,
      photo_url: photoUrl,
      file_id: existing?.file_id || "",
      photo_order: String(index + 1),
      is_primary: index === 0 ? "TRUE" : "FALSE",
      status: "active",
      created_at: existing?.created_at || today,
      updated_at: today,
    };
    const row = headers.map((header) => (valuesByHeader[header] !== undefined ? valuesByHeader[header] : ""));
    if (existingRowIndex !== undefined) {
      await updateSheetValues(`user_photos!A${existingRowIndex + 2}:${columnName(headers.length)}${existingRowIndex + 2}`, [row]);
    } else {
      await appendSheetValues(`user_photos!A:${columnName(headers.length)}`, [row]);
    }
  }

  for (let index = photos.length; index < existingIndexes.length; index += 1) {
    const { photo, index: rowIndex } = existingIndexes[index];
    const valuesByHeader = {
      ...photo,
      status: "deleted",
      updated_at: today,
    };
    const row = headers.map((header) => (valuesByHeader[header] !== undefined ? valuesByHeader[header] : ""));
    await updateSheetValues(`user_photos!A${rowIndex + 2}:${columnName(headers.length)}${rowIndex + 2}`, [row]);
  }

  return photos.length;
}

async function saveMeetingPlace({ placeId, userId, profile, today, existingPlaceId }) {
  const current = await getSheetValues("meeting_places!A1:Z1000");
  const headers = await ensureMeetingPlaceHeaders(current.values?.[0] || []);
  const places = rowsToObjects(current.values);
  const existingIndex = places.findIndex(
    (place) => place.place_id === existingPlaceId || place.place_id === placeId || place.owner_user_id === userId,
  );
  const existing = existingIndex >= 0 ? places[existingIndex] : null;
  const placeName = profile.meetingPlaceName || firstLineValue(profile.meetingArea, "地點") || profile.meetingArea || "";
  const mapUrl = profile.meetingPlaceUrl || firstLineValue(profile.meetingArea, "Google Maps") || "";
  const valuesByHeader = {
    place_id: placeId,
    place_name: placeName,
    address: existing?.address || "",
    city: profile.city || existing?.city || "",
    district: profile.district || existing?.district || "",
    map_url: mapUrl,
    lat: profile.meetingLat || firstLineCoords(profile.meetingArea)?.lat || existing?.lat || "",
    lng: profile.meetingLng || firstLineCoords(profile.meetingArea)?.lng || existing?.lng || "",
    google_place_id: existing?.google_place_id || "",
    category: existing?.category || "coffee",
    is_public: existing?.is_public || "FALSE",
    owner_user_id: userId,
    created_at: existing?.created_at || today,
    updated_at: today,
    raw_input: profile.meetingArea || "",
  };
  const row = headers.map((header) => (valuesByHeader[header] !== undefined ? valuesByHeader[header] : ""));
  if (existing) {
    await updateSheetValues(`meeting_places!A${existingIndex + 2}:${columnName(headers.length)}${existingIndex + 2}`, [row]);
    return;
  }
  await appendSheetValues(`meeting_places!A:${columnName(headers.length)}`, [row]);
}

function firstLineValue(value, label) {
  const match = String(value || "").match(new RegExp(`^${label}：(.+)$`, "m"));
  return match ? match[1].trim() : "";
}

function firstLineCoords(value) {
  const match = String(value || "").match(/^座標：\s*(-?\d+(?:\.\d+)?),\s*(-?\d+(?:\.\d+)?)/m);
  return match ? { lat: match[1], lng: match[2] } : null;
}

async function ensureMeetingPlaceHeaders(currentHeaders) {
  const requiredHeaders = [
    "place_id",
    "place_name",
    "address",
    "city",
    "district",
    "map_url",
    "lat",
    "lng",
    "google_place_id",
    "category",
    "is_public",
    "owner_user_id",
    "created_at",
    "updated_at",
    "raw_input",
  ];
  const headers = currentHeaders.filter(Boolean);
  const mergedHeaders = headers.length
    ? headers.concat(requiredHeaders.filter((header) => !headers.includes(header)))
    : requiredHeaders;
  if (mergedHeaders.length !== headers.length) {
    await updateSheetValues(`meeting_places!A1:${columnName(mergedHeaders.length)}1`, [mergedHeaders]);
  }
  return mergedHeaders;
}

async function ensurePhotoHeaders(currentHeaders) {
  const requiredHeaders = [
    "photo_id",
    "user_id",
    "photo_url",
    "file_id",
    "photo_order",
    "is_primary",
    "status",
    "created_at",
    "updated_at",
  ];
  const headers = currentHeaders.filter(Boolean);
  const mergedHeaders = headers.length
    ? headers.concat(requiredHeaders.filter((header) => !headers.includes(header)))
    : requiredHeaders;
  if (mergedHeaders.length !== headers.length) {
    await updateSheetValues(`user_photos!A1:${columnName(mergedHeaders.length)}1`, [mergedHeaders]);
  }
  return mergedHeaders;
}

async function ensureUserHeaders(currentHeaders) {
  const requiredHeaders = [
    "user_id",
    "nickname",
    "gender",
    "age",
    "city",
    "district",
    "email",
    "occupation",
    "education",
    "school",
    "relationship_status",
    "has_children",
    "coffee_goal",
    "smoking",
    "drinking",
    "intro",
    "interest_keywords",
    "available_times",
    "meeting_place_id",
    "lock_until",
    "status",
    "created_at",
    "updated_at",
    "notes",
    "birth_year",
    "birth_month",
    "birth_day",
  ];
  const headers = currentHeaders.filter(Boolean);
  const mergedHeaders = headers.length
    ? headers.concat(requiredHeaders.filter((header) => !headers.includes(header)))
    : requiredHeaders;
  if (mergedHeaders.length !== headers.length) {
    await updateSheetValues(`users!A1:${columnName(mergedHeaders.length)}1`, [mergedHeaders]);
  }
  return mergedHeaders;
}

function columnName(index) {
  let name = "";
  while (index > 0) {
    index -= 1;
    name = String.fromCharCode(65 + (index % 26)) + name;
    index = Math.floor(index / 26);
  }
  return name || "A";
}

function sendJson(res, status, data) {
  res.writeHead(status, {
    "content-type": "application/json; charset=utf-8",
    "access-control-allow-origin": "*",
  });
  res.end(JSON.stringify(data, null, 2));
}

function readRequestBody(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", (chunk) => {
      body += chunk;
      if (body.length > 25_000_000) {
        reject(new Error("Request body too large"));
        req.destroy();
      }
    });
    req.on("end", () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch {
        reject(new Error("Invalid JSON body"));
      }
    });
    req.on("error", reject);
  });
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
    if (req.method === "OPTIONS") {
      sendJson(res, 200, { ok: true });
      return;
    }

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

    if (url.pathname === "/api/users" && req.method === "POST") {
      const profile = await readRequestBody(req);
      const result = await createUserFromProfile(profile);
      sendJson(res, 200, { ok: true, ...result });
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
