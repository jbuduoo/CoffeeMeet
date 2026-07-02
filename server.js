const http = require("http");
const fs = require("fs");
const path = require("path");
const { appendSheetValues, batchUpdateSpreadsheet, getSpreadsheetMeta, getSheetValues, updateSheetValues } = require("./sheets-client");

const port = Number(process.env.PORT || 3000);
const rootDir = __dirname;
const visitSheetName = "上來紀錄";
const visitHeaders = ["上來時間", "user_id", "姓名/暱稱", "email", "來源/活動", "備註"];
const visitStatsSheetName = "人員統計";
const visitStatsHeaders = [
  "user_id",
  "姓名/暱稱",
  "email",
  "狀態",
  "會員建立日",
  "會員更新日",
  "上來次數",
  "最近上來",
  "首次上來",
  "近7天",
  "近30天",
  "活躍狀態",
  "備註",
];

function extractGoogleMapsUrl(input = "") {
  const verifiedUrlMatch = String(input || "").match(/Google Maps：(\S+)/);
  return verifiedUrlMatch ? verifiedUrlMatch[1].trim() : String(input || "").trim();
}

function parseGoogleMapsPlaceUrl(source, originalInput = "") {
  const decodedPath = decodeURIComponent(source.replace(/\+/g, " "));
  const placeMatch = decodedPath.match(/\/place\/([^/@?]+)/);
  const verifiedNameMatch = String(originalInput || "").match(/地點：([^｜\n]+)/);
  const rawName = verifiedNameMatch ? verifiedNameMatch[1].trim() : placeMatch ? placeMatch[1].trim() : "";
  const verifiedCoordsMatch = String(originalInput || "").match(/座標：\s*(-?\d+(?:\.\d+)?),\s*(-?\d+(?:\.\d+)?)/);
  const preciseMatch = source.match(/!3d(-?\d+(?:\.\d+)?)!4d(-?\d+(?:\.\d+)?)/);
  const atMatch = source.match(/@(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/);
  const coords = verifiedCoordsMatch || preciseMatch || atMatch;
  if (!coords) return null;
  return {
    name: rawName || "Google Maps 地點",
    url: source,
    lat: coords[1],
    lng: coords[2],
  };
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
    "路易莎",
    "louisa",
    "咖啡",
    "coffee",
    "cafe",
    "café",
  ].some((keyword) => normalized.includes(keyword));
}

function assertPublicMeetupPlace(place) {
  if (!hasPublicMeetupKeyword(place && place.name)) {
    throw new Error("請選擇店名含有全家、7-11、星巴克、路易莎或咖啡字樣的公開店點。");
  }
}

async function expandGoogleMapsShortUrl(input) {
  const source = extractGoogleMapsUrl(input);
  const urlParts = parseUrlParts(source);
  if (!urlParts) {
    const error = new Error("請貼上有效的 Google Maps 地點網址");
    error.debug = {
      step: "parse-url",
      source,
      reason: "Could not parse the submitted URL string.",
    };
    throw error;
  }
  if (!/google\.[^/]+\/maps|maps\.app\.goo\.gl/.test(urlParts.hostname + urlParts.pathname)) {
    throw new Error("請貼上 Google Maps 地點網址");
  }
  if (!urlParts.hostname.includes("maps.app.goo.gl")) return source;

  const response = await fetch(source, { method: "HEAD", redirect: "manual" });
  const location = response.headers.get("location");
  if (!location) throw new Error("無法展開 Google Maps 短網址，請改貼完整網址");
  return location;
}

function parseUrlParts(value) {
  const match = String(value || "").trim().match(/^https?:\/\/([^/?#]+)([^?#]*)/i);
  if (!match) return null;
  return {
    hostname: match[1].toLowerCase(),
    pathname: match[2] || "/",
  };
}

function isGoogleMapsInput(input) {
  const source = extractGoogleMapsUrl(input);
  const urlParts = parseUrlParts(source);
  return Boolean(urlParts && /google\.[^/]+\/maps|maps\.app\.goo\.gl/.test(urlParts.hostname + urlParts.pathname));
}

async function searchGooglePlacesByText(query) {
  const normalizedQuery = String(query || "").trim();
  if (!normalizedQuery) return [];
  const apiKey = process.env.GOOGLE_MAPS_API_KEY;
  if (!apiKey) throw new Error("Missing GOOGLE_MAPS_API_KEY for road lookup");
  const response = await fetch("https://places.googleapis.com/v1/places:searchText", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "X-Goog-Api-Key": apiKey,
      "X-Goog-FieldMask": "places.id,places.displayName,places.formattedAddress,places.location,places.googleMapsUri",
    },
    body: JSON.stringify({
      textQuery: normalizedQuery,
      languageCode: "zh-TW",
      regionCode: "TW",
      maxResultCount: 3,
    }),
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok || data.error) throw new Error(data?.error?.message || "Google Places lookup failed");
  return (data.places || []).slice(0, 3).map((place) => {
    const location = place.location || {};
    return {
      name: place.displayName?.text || place.formattedAddress || "road",
      url: place.googleMapsUri || "",
      lat: location.latitude || "",
      lng: location.longitude || "",
      placeId: place.id || "",
      address: place.formattedAddress || "",
    };
  });
}

async function verifyGoogleMapsPlace(payload) {
  const input = String(payload?.url || payload?.input || payload?.keyword || payload?.query || "").trim();
  const city = String(payload?.city || "").trim();
  const district = String(payload?.district || "").trim();
  const mode = String(payload?.mode || "").trim();
  const query = String(payload?.query || [input, city, district].filter(Boolean).join(" ")).trim();
  if (!input) throw new Error("Missing place input");
  if (!isGoogleMapsInput(input)) {
    const places = await searchGooglePlacesByText(query || input);
    if (mode === "road-geocode") return { ok: true, place: places[0] || null, results: places };
    const allowedPlaces = places.filter((place) => hasPublicMeetupKeyword(place.name));
    return { ok: true, place: allowedPlaces[0] || null, results: allowedPlaces };
  }
  const expandedUrl = await expandGoogleMapsShortUrl(input);
  const place = parseGoogleMapsPlaceUrl(expandedUrl, input);
  if (!place) throw new Error("請貼上有效的 Google Maps 地點網址");
  return { place };
}

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

function normalizePhotoUrl(fileId = "", photoUrl = "") {
  const driveFileId = fileId || extractDriveFileId(photoUrl);
  if (driveFileId) return `https://lh3.googleusercontent.com/d/${driveFileId}=w1200`;
  return photoUrl || "";
}

async function getAppData() {
  const [usersSheet, placesSheet, photosSheet, invitesSheet] = await Promise.all([
    getSheetValues("users!A1:AZ1000"),
    getSheetValues("meeting_places!A1:AZ1000"),
    getSheetValues("user_photos!A1:AZ1000"),
    getSheetValues("invites!A1:AZ1000"),
  ]);

  const users = rowsToObjects(usersSheet.values);
  const places = rowsToObjects(placesSheet.values);
  const photos = rowsToObjects(photosSheet.values).filter((photo) => photo.status !== "deleted");
  const inviteRows = rowsToObjects(invitesSheet.values);
  const placesById = Object.fromEntries(places.map((place) => [place.place_id, place]));
  const placesByOwnerUserId = Object.fromEntries(places.map((place) => [place.owner_user_id, place]).filter(([userId]) => userId));
  const photosByUserId = photos.reduce((groups, photo) => {
    if (!photo.user_id || (!photo.file_id && !photo.photo_url)) return groups;
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
        socialStyleKeywords: user.social_style_keywords || user.social_styles || "",
        socialStyles: user.social_style_keywords || user.social_styles || "",
        photo: normalizePhotoUrl(primaryPhoto.file_id, primaryPhoto.photo_url),
        photos: userPhotos.map((photo) => normalizePhotoUrl(photo.file_id, photo.photo_url)).filter(Boolean),
      };
    });

  const fallbackPlace = (placeId) => placesById[placeId]?.place_name || "";
  const invites = inviteRows.map((invite, index) => ({
    id: invite.invite_id || `sheet-invite-${index + 1}`,
    inviteId: invite.invite_id || "",
    rowIndex: index + 1,
    candidateId: invite.receiver_user_id || invite.sender_user_id,
    senderUserId: invite.sender_user_id || "",
    receiverUserId: invite.receiver_user_id || "",
    status: statusToInviteStatus(invite.status),
    feedbackScore: invite.feedback_score || "",
    maleFeedbackScore: invite.male_feedback_score || "",
    femaleFeedbackScore: invite.female_feedback_score || "",
    createdAt: invite.created_at || "",
    updatedAt: invite.updated_at || "",
    note: stripAddressLines(invite.note) || [
      invite.accepted_at ? `已確認：${invite.accepted_at}` : "",
      invite.selected_times ? `可選時段：${invite.selected_times}` : "",
      `地點：${invite.place_name || fallbackPlace(invite.receiver_place_id || invite.sender_place_id)}`,
    ]
      .filter(Boolean)
      .join("；") || "從試算表載入的邀約",
  }));

  return { candidates, invites };
}

const appDataCacheTtlMs = 30 * 1000;
let appDataCache = {
  data: null,
  expiresAt: 0,
};
let appDataCachePromise = null;

function clearAppDataCache() {
  appDataCache = {
    data: null,
    expiresAt: 0,
  };
  appDataCachePromise = null;
}

async function getCachedAppData() {
  const now = Date.now();
  if (appDataCache.data && now < appDataCache.expiresAt) {
    return {
      ...appDataCache.data,
      cached: true,
      cacheExpiresAt: new Date(appDataCache.expiresAt).toISOString(),
    };
  }

  if (!appDataCachePromise) {
    appDataCachePromise = getAppData()
      .then((freshData) => {
        appDataCache = {
          data: freshData,
          expiresAt: Date.now() + appDataCacheTtlMs,
        };
        return freshData;
      })
      .finally(() => {
        appDataCachePromise = null;
      });
  }

  const freshData = await appDataCachePromise;
  return {
    ...freshData,
    cached: false,
    cacheExpiresAt: new Date(appDataCache.expiresAt).toISOString(),
  };
}

function todayString() {
  return new Date().toISOString().slice(0, 10);
}

function calculateAge(birthYear, today = new Date()) {
  const year = Number(birthYear);
  if (!year) return "";
  return String(today.getFullYear() - year);
}

function normalizeEmail(value) {
  return String(value || "")
    .trim()
    .split(/[?#&]/)[0]
    .trim()
    .toLowerCase();
}

function isValidEmail(value) {
  return /^[^\s@?&=]+@[^\s@?&=]+\.[^\s@?&=]+$/.test(value);
}

function hasValidCoords(place = {}) {
  const lat = Number(place.lat);
  const lng = Number(place.lng);
  return Number.isFinite(lat) && Number.isFinite(lng) && Math.abs(lat) <= 90 && Math.abs(lng) <= 180;
}

function isInviteReadyUser(user = {}, photos = [], place = {}) {
  const basicFields = [
    user.nickname,
    user.gender,
    user.age || user.birth_year,
    user.city,
    user.district,
    user.coffee_goal,
    user.intro,
  ];
  return (
    basicFields.every((value) => String(value || "").trim()) &&
    photos.some((photo) => String(photo.status || "active") !== "deleted" && (photo.file_id || photo.photo_url)) &&
    Boolean(String(user.available_times || "").trim()) &&
    Boolean(String(place.place_name || user.meeting_place_id || "").trim()) &&
    hasValidCoords(place)
  );
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
  const email = normalizeEmail(profile.email);
  if (!email) throw new Error("Email is required");
  if (!isValidEmail(email)) throw new Error("Invalid email format");

  const current = await getSheetValues("users!A1:AZ1000");
  const headers = await ensureUserHeaders(current.values?.[0] || []);
  const users = rowsToObjects(current.values);
  const existingIndex = users.findIndex((user) => normalizeEmail(user.email) === email);
  const existing = existingIndex >= 0 ? users[existingIndex] : null;

  const stamp = Date.now().toString(36);
  const userId = existing?.user_id || `${profile.gender === "女性" ? "girl" : "user"}-${slugFromEmail(email)}-${stamp}`;
  const placeId = existing?.meeting_place_id || `place-${slugFromEmail(email)}-${stamp}`;
  const today = todayString();
  const age = calculateAge(profile.birthYear);

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
    birth_month: "",
    birth_day: "",
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
    social_style_keywords: profile.socialStyleKeywords || profile.socialStyles || "",
    social_styles: profile.socialStyleKeywords || profile.socialStyles || "",
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

async function createInvite(invite) {
  const senderUserId = String(invite.senderUserId || invite.sender_user_id || "").trim();
  const receiverUserId = String(invite.receiverUserId || invite.receiver_user_id || "").trim();
  if (!senderUserId || !receiverUserId) throw new Error("senderUserId and receiverUserId are required");

  const [userValues, photoValues, placeValues] = await Promise.all([
    getSheetValues("users!A1:AZ1000"),
    getSheetValues("user_photos!A1:AZ1000"),
    getSheetValues("meeting_places!A1:AZ1000"),
  ]);
  const users = rowsToObjects(userValues.values);
  const photos = rowsToObjects(photoValues.values);
  const places = rowsToObjects(placeValues.values);
  const placesById = Object.fromEntries(places.map((place) => [place.place_id, place]));
  const placesByOwnerUserId = Object.fromEntries(places.map((place) => [place.owner_user_id, place]));
  const photosByUserId = photos.reduce((groups, photo) => {
    groups[photo.user_id] ||= [];
    groups[photo.user_id].push(photo);
    return groups;
  }, {});
  const sender = users.find((user) => user.user_id === senderUserId);
  const receiver = users.find((user) => user.user_id === receiverUserId);
  const senderPlace = placesById[sender?.meeting_place_id] || placesByOwnerUserId[senderUserId] || {};
  const receiverPlace = placesById[receiver?.meeting_place_id] || placesByOwnerUserId[receiverUserId] || {};
  if (!isInviteReadyUser(sender, photosByUserId[senderUserId] || [], senderPlace)) {
    throw new Error("完成名片與咖啡地點定位後，才能開始邀約。");
  }
  if (!isInviteReadyUser(receiver, photosByUserId[receiverUserId] || [], receiverPlace)) {
    throw new Error("對方名片尚未完成，暫時不能接受邀約。");
  }

  const current = await getSheetValues("invites!A1:AZ1000");
  const headers = await ensureInviteHeaders(current.values?.[0] || []);
  const today = todayString();
  const note = stripAddressLines(invite.note).trim();
  const selectedTimes = Array.isArray(invite.selectedTimes)
    ? invite.selectedTimes.join("、")
    : String(invite.selected_times || invite.selectedTimes || detailFromNote(note, "可選時段") || "").trim();
  const placeName = String(invite.placeName || invite.place_name || detailFromNote(note, "地點") || "").trim();
  const feedbackScore = String(invite.feedbackScore || invite.feedback_score || detailFromNote(note, "問卷分數").replace(/分$/, "") || "").trim();
  const feedbackGender = String(invite.feedbackGender || invite.feedback_gender || "").trim();
  const maleFeedbackScore = String(invite.maleFeedbackScore || invite.male_feedback_score || (feedbackGender === "男性" ? feedbackScore : "")).trim();
  const femaleFeedbackScore = String(invite.femaleFeedbackScore || invite.female_feedback_score || (feedbackGender === "女性" ? feedbackScore : "")).trim();
  const hasParticipantFeedbackScore = Boolean(maleFeedbackScore || femaleFeedbackScore);
  const suppliedInviteId = String(invite.inviteId || invite.invite_id || invite.id || "").trim();
  const inviteId = suppliedInviteId || `invite-${Date.now().toString(36)}`;
  const rows = rowsToObjects(current.values);
  const requestedStatus = statusToInviteStatus(invite.status || "");
  let existingIndex = suppliedInviteId
    ? rows.findIndex((row) => String(row.invite_id || "").trim() === suppliedInviteId)
    : -1;
  if (existingIndex < 0) {
    existingIndex = rows.findIndex((row) =>
      ((row.sender_user_id === senderUserId && row.receiver_user_id === receiverUserId) ||
        (row.sender_user_id === receiverUserId && row.receiver_user_id === senderUserId)) &&
      ["sent", "pending", "incoming", "confirmed"].includes(statusToInviteStatus(row.status)),
    );
  }
  if (existingIndex >= 0 && ["confirmed", "cancelled", "done"].includes(requestedStatus)) {
    const existing = rows[existingIndex];
    const acceptedTime = String(invite.acceptedTime || invite.accepted_at || detailFromNote(note, "已確認") || "").trim();
    const valuesByHeader = {
      ...existing,
      status: requestedStatus,
      note: note || stripAddressLines(requestedStatus === "confirmed" ? `已確認：${acceptedTime}；${existing.note || ""}` : existing.note || ""),
      feedback_score: hasParticipantFeedbackScore ? existing.feedback_score || "" : feedbackScore || existing.feedback_score || "",
      male_feedback_score: maleFeedbackScore || existing.male_feedback_score || "",
      female_feedback_score: femaleFeedbackScore || existing.female_feedback_score || "",
      accepted_at: acceptedTime || existing.accepted_at || "",
      updated_at: today,
    };
    const row = headers.map((header) => (valuesByHeader[header] !== undefined ? valuesByHeader[header] : ""));
    await updateSheetValues(`invites!A${existingIndex + 2}:${columnName(headers.length)}${existingIndex + 2}`, [row]);
    return {
      inviteId: existing.invite_id || inviteId,
      updated: true,
      rowIndex: existingIndex + 1,
      beforeStatus: existing.status || "",
      afterStatus: requestedStatus,
    };
  }
  if (existingIndex < 0 && ["confirmed", "cancelled", "done"].includes(requestedStatus)) {
    throw new Error("Invite update target not found");
  }
  if (existingIndex >= 0) {
    return { inviteId: rows[existingIndex].invite_id || inviteId, duplicate: true };
  }

  const valuesByHeader = {
    invite_id: inviteId,
    sender_user_id: senderUserId,
    receiver_user_id: receiverUserId,
    status: "sent",
    selected_times: selectedTimes,
    place_name: placeName,
    note: note || `可選時段：${selectedTimes}；地點：${placeName}`,
    feedback_score: hasParticipantFeedbackScore ? "" : feedbackScore,
    male_feedback_score: maleFeedbackScore,
    female_feedback_score: femaleFeedbackScore,
    created_at: today,
    updated_at: today,
    accepted_at: "",
  };
  const row = headers.map((header) => (valuesByHeader[header] !== undefined ? valuesByHeader[header] : ""));
  await appendSheetValues(`invites!A:${columnName(headers.length)}`, [row]);
  return { inviteId };
}

async function ensureVisitSheet() {
  const meta = await getSpreadsheetMeta();
  const missingSheets = [visitSheetName, visitStatsSheetName].filter((sheetName) => !meta.sheets.includes(sheetName));
  if (missingSheets.length) {
    await batchUpdateSpreadsheet(missingSheets.map((sheetName) => ({
      addSheet: {
        properties: {
          title: sheetName,
          gridProperties: {
            rowCount: sheetName === visitSheetName ? 1000 : 500,
            columnCount: sheetName === visitSheetName ? visitHeaders.length : visitStatsHeaders.length,
            frozenRowCount: 1,
          },
        },
      },
    })));
  }
  await updateSheetValues(`${visitSheetName}!A1:${columnName(visitHeaders.length)}1`, [visitHeaders]);
  await updateSheetValues(`${visitStatsSheetName}!A1:${columnName(visitStatsHeaders.length)}1`, [visitStatsHeaders]);
}

function visitStatsRow(payload, rowNumber) {
  const userId = String(payload.userId || payload.user_id || "").trim();
  const email = normalizeEmail(payload.email);
  const name = String(payload.name || payload.nickname || email || userId).trim();
  const logSheet = `'${visitSheetName.replace(/'/g, "''")}'`;
  return [
    userId,
    name,
    email,
    String(payload.status || "").trim(),
    String(payload.createdAt || payload.created_at || "").trim(),
    String(payload.updatedAt || payload.updated_at || "").trim(),
    `=COUNTIF(${logSheet}!B:B,A${rowNumber})`,
    `=IF(G${rowNumber}=0,"",MAXIFS(${logSheet}!A:A,${logSheet}!B:B,A${rowNumber}))`,
    `=IF(G${rowNumber}=0,"",MINIFS(${logSheet}!A:A,${logSheet}!B:B,A${rowNumber}))`,
    `=COUNTIFS(${logSheet}!B:B,A${rowNumber},${logSheet}!A:A,">="&TODAY()-7)`,
    `=COUNTIFS(${logSheet}!B:B,A${rowNumber},${logSheet}!A:A,">="&TODAY()-30)`,
    `=IF(G${rowNumber}=0,"未記錄",IF(H${rowNumber}>=TODAY()-7,"活躍",IF(H${rowNumber}>=TODAY()-30,"普通","沉寂")))`,
    "",
  ];
}

async function upsertVisitStatsRow(payload = {}) {
  const userId = String(payload.userId || payload.user_id || "").trim();
  const email = normalizeEmail(payload.email);
  if (!userId && !email) return;

  const current = await getSheetValues(`${visitStatsSheetName}!A1:M2000`);
  const rows = current.values || [];
  const bodyRows = rows.slice(1);
  const existingIndex = bodyRows.findIndex((row) =>
    (userId && String(row[0] || "").trim() === userId) ||
    (email && normalizeEmail(row[2]) === email),
  );
  const rowNumber = existingIndex >= 0 ? existingIndex + 2 : bodyRows.length + 2;
  const row = visitStatsRow(payload, rowNumber);

  if (existingIndex >= 0) {
    await updateSheetValues(`${visitStatsSheetName}!A${rowNumber}:M${rowNumber}`, [row]);
    return;
  }
  await appendSheetValues(`${visitStatsSheetName}!A:M`, [row]);
}

async function recordVisit(payload = {}) {
  const userId = String(payload.userId || payload.user_id || "").trim();
  const email = normalizeEmail(payload.email);
  const name = String(payload.name || payload.nickname || "").trim();
  if (!userId && !email) throw new Error("userId or email is required");

  await ensureVisitSheet();
  await appendSheetValues(`${visitSheetName}!A:${columnName(visitHeaders.length)}`, [[
    new Date().toISOString(),
    userId,
    name,
    email,
    String(payload.source || "login").trim() || "login",
    String(payload.note || "").trim(),
  ]]);
  await upsertVisitStatsRow(payload);

  return { recorded: true };
}

function detailFromNote(note, label) {
  const match = String(note || "").match(new RegExp(`${label}：([^；]+)`));
  return match ? match[1].trim() : "";
}
async function saveUserPhotos({ userId, profile, today }) {
  const photos = Array.isArray(profile.photos) ? profile.photos.filter(Boolean).slice(0, 3) : [];
  const current = await getSheetValues("user_photos!A1:AZ1000");
  const headers = await ensurePhotoHeaders(current.values?.[0] || []);
  const rows = rowsToObjects(current.values);
  const existingIndexes = rows
    .map((photo, index) => ({ photo, index }))
    .filter(({ photo }) => photo.user_id === userId);

  for (let index = 0; index < photos.length; index += 1) {
    const existing = existingIndexes[index]?.photo;
    const existingRowIndex = existingIndexes[index]?.index;
    const photoValue = photos[index];
    const fileId = extractDriveFileId(photoValue) || existing?.file_id || "";
    const valuesByHeader = {
      photo_id: existing?.photo_id || `photo-${userId}-${index + 1}`,
      user_id: userId,
      photo_url: "",
      file_id: fileId,
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

function extractDriveFileId(value) {
  const text = String(value || "").trim();
  const ucMatch = text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (ucMatch) return ucMatch[1];
  const fileMatch = text.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (fileMatch) return fileMatch[1];
  return /^[a-zA-Z0-9_-]{20,}$/.test(text) ? text : "";
}

async function saveMeetingPlace({ placeId, userId, profile, today, existingPlaceId }) {
  const current = await getSheetValues("meeting_places!A1:AZ1000");
  const headers = await ensureMeetingPlaceHeaders(current.values?.[0] || []);
  const places = rowsToObjects(current.values);
  const existingIndex = places.findIndex(
    (place) => place.place_id === existingPlaceId || place.place_id === placeId || place.owner_user_id === userId,
  );
  const existing = existingIndex >= 0 ? places[existingIndex] : null;
  const sanitizedMeetingArea = stripAddressLines(profile.meetingArea);
  const placeName = profile.meetingPlaceName || firstLineValue(sanitizedMeetingArea, "路段") || firstLineValue(sanitizedMeetingArea, "地點") || sanitizedMeetingArea || "";
  const mapUrl = profile.meetingPlaceUrl || firstLineValue(profile.meetingArea, "Google Maps") || "";
  const valuesByHeader = {
    place_id: placeId,
    place_name: placeName,
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
    raw_input: sanitizedMeetingArea,
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

function stripAddressLines(value) {
  return String(value || "")
    .split(/\r?\n/)
    .filter((line) => !/^地址：/.test(line.trim()))
    .join("\n");
}

function firstLineCoords(value) {
  const match = String(value || "").match(/^座標：\s*(-?\d+(?:\.\d+)?),\s*(-?\d+(?:\.\d+)?)/m);
  return match ? { lat: match[1], lng: match[2] } : null;
}

async function ensureMeetingPlaceHeaders(currentHeaders) {
  const requiredHeaders = [
    "place_id",
    "place_name",
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
    "social_style_keywords",
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

async function ensureInviteHeaders(currentHeaders) {
  const requiredHeaders = [
    "invite_id",
    "sender_user_id",
    "receiver_user_id",
    "status",
    "selected_times",
    "place_name",
    "note",
    "feedback_score",
    "male_feedback_score",
    "female_feedback_score",
    "created_at",
    "updated_at",
    "accepted_at",
  ];
  const headers = currentHeaders.filter(Boolean);
  const mergedHeaders = headers.length
    ? headers.concat(requiredHeaders.filter((header) => !headers.includes(header)))
    : requiredHeaders;
  if (mergedHeaders.length !== headers.length) {
    await updateSheetValues(`invites!A1:${columnName(mergedHeaders.length)}1`, [mergedHeaders]);
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
      const data = await getCachedAppData();
      sendJson(res, 200, { ok: true, ...data });
      return;
    }

    if (["/api/places-verify", "/api/places%2Fverify", "/api/places/verify"].includes(url.pathname) && req.method === "POST") {
      const payload = await readRequestBody(req);
      const result = await verifyGoogleMapsPlace(payload);
      sendJson(res, 200, { ok: true, ...result });
      return;
    }

    if (url.pathname === "/api/users" && req.method === "POST") {
      const profile = await readRequestBody(req);
      const result = await createUserFromProfile(profile);
      clearAppDataCache();
      sendJson(res, 200, { ok: true, ...result });
      return;
    }

    if (url.pathname === "/api/invites" && req.method === "POST") {
      const invite = await readRequestBody(req);
      const result = await createInvite(invite);
      clearAppDataCache();
      sendJson(res, 200, { ok: true, ...result });
      return;
    }

    if (url.pathname === "/api/visits" && req.method === "POST") {
      const payload = await readRequestBody(req);
      const result = await recordVisit(payload);
      sendJson(res, 200, { ok: true, ...result });
      return;
    }

    sendJson(res, 404, {
      ok: false,
      error: "API not found",
      debug: {
        method: req.method,
        pathname: url.pathname,
        search: url.search,
        availablePaths: ["/api/app-data", "/api/sheets", "/api/sheets/values", "/api/places-verify", "/api/users", "/api/invites", "/api/visits"],
      },
    });
  } catch (error) {
    sendJson(res, 500, { ok: false, error: error.message, debug: error.debug || null, stack: error.stack || "" });
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
