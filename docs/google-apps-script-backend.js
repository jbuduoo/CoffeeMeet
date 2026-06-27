const SPREADSHEET_ID = "1JzQwlNWQphrHfUCcpQl3pC6jHLXN6ykwvJKSs5Uaxsg";
const BOY_FOLDER_ID = "12M91XLYJ7seBNhrsGL2JZNJZlSGlL9ri";
const GIRL_FOLDER_ID = "1r_6bPvf-R5lHM92g379al71G8zzmj-of";
const APP_URL = "https://jbuduoo.github.io/CoffeeMeet/app.html";
const EMAIL_FALLBACK_TO = "jbuduoo@gmail.com";

function doGet(event) {
  const params = (event && event.parameter) || {};
  const action = params.action || "";
  if (action === "app-data") return jsonResponse(getAppData());
  if (action === "sheets") return jsonResponse(getSheetMeta());
  if (action === "version") return jsonResponse(getVersion());
  if (action === "photo-storage") return jsonResponse(getPhotoStorageMeta());
  if (action === "publish-seed-photos") return jsonResponse(publishSeedPhotos(params.token || ""));
  if (action === "clear-photo-urls") return jsonResponse(clearPhotoUrls(params.token || ""));
  return jsonResponse(debugError("Unknown action", action, "GET"));
}

function doPost(event) {
  try {
    const params = (event && event.parameter) || {};
    const action = params.action || "";
    const body = JSON.parse((event && event.postData && event.postData.contents) || "{}");
    if (action === "places-verify" || action === "places/verify") return jsonResponse(verifyGoogleMapsPlace(body));
    if (action === "users") return jsonResponse(saveUserProfile(body));
    if (action === "invites") return jsonResponse(saveInvite(body));
    return jsonResponse(debugError("Unknown action", action, "POST", body));
  } catch (error) {
    return jsonResponse({
      ok: false,
      error: error.message || String(error),
      debug: error.debug || null,
      stack: error.stack || "",
    });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function verifyGoogleMapsPlace(payload) {
  const input = String((payload && payload.url) || "").trim();
  const debug = {
    action: "places-verify",
    input: input,
    startedAt: new Date().toISOString(),
  };
  const expandedUrl = expandGoogleMapsShortUrl(input);
  debug.expandedUrl = expandedUrl;
  const place = parseGoogleMapsPlaceUrl(expandedUrl, input);
  debug.parsedPlace = place;
  if (!place) {
    const error = new Error("請貼上有效的 Google Maps 地點網址");
    error.debug = debug;
    throw error;
  }
  return { ok: true, place, debug: debug };
}

function expandGoogleMapsShortUrl(input) {
  const source = extractGoogleMapsUrl(input);
  const urlParts = parseUrlParts(source);
  if (!urlParts) {
    const error = new Error("請貼上有效的 Google Maps 地點網址");
    error.debug = {
      step: "parse-url",
      source: source,
      reason: "Apps Script could not parse the submitted URL string.",
    };
    throw error;
  }
  if (!/google\.[^/]+\/maps|maps\.app\.goo\.gl/.test(urlParts.hostname + urlParts.pathname)) {
    throw new Error("請貼上 Google Maps 地點網址");
  }
  if (urlParts.hostname.indexOf("maps.app.goo.gl") < 0) return source;

  const response = UrlFetchApp.fetch(source, {
    method: "get",
    followRedirects: false,
    muteHttpExceptions: true,
    headers: {
      "User-Agent": "Mozilla/5.0",
    },
  });
  const headers = response.getAllHeaders();
  const location = headers.Location || headers.location;
  if (!location) {
    const error = new Error("無法展開 Google Maps 短網址，請改貼完整網址");
    error.debug = {
      source: source,
      responseCode: response.getResponseCode(),
      headers: headers,
      bodyPreview: response.getContentText().slice(0, 500),
    };
    throw error;
  }
  return location;
}

function debugError(message, action, method, body) {
  return {
    ok: false,
    error: message,
    debug: {
      method: method,
      receivedAction: action,
      body: body || null,
      availableActions: ["app-data", "sheets", "version", "photo-storage", "publish-seed-photos", "clear-photo-urls", "places-verify", "users", "invites"],
      deployedAtHint: "If places-verify is missing here, redeploy the latest Apps Script code.",
    },
  };
}

function parseGoogleMapsPlaceUrl(source, originalInput) {
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

function extractGoogleMapsUrl(input) {
  const verifiedUrlMatch = String(input || "").match(/Google Maps：(\S+)/);
  return verifiedUrlMatch ? verifiedUrlMatch[1].trim() : String(input || "").trim();
}

function parseUrlParts(value) {
  const match = String(value || "").trim().match(/^https?:\/\/([^/?#]+)([^?#]*)/i);
  if (!match) return null;
  return {
    hostname: match[1].toLowerCase(),
    pathname: match[2] || "/",
  };
}

function sheetRows(sheetName) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift() || [];
  return values
    .filter((row) => row.some((cell) => String(cell || "").trim()))
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] || ""])));
}

function getSheetMeta() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    ok: true,
    title: spreadsheet.getName(),
    sheets: spreadsheet.getSheets().map((sheet) => sheet.getName()),
  };
}

function getAppData() {
  const users = sheetRows("users");
  const places = sheetRows("meeting_places");
  const photos = sheetRows("user_photos").filter((photo) => photo.status !== "deleted");
  const invites = sheetRows("invites");
  const placesById = Object.fromEntries(places.map((place) => [place.place_id, place]));
  const placesByOwnerUserId = Object.fromEntries(
    places.map((place) => [place.owner_user_id, place]).filter(([userId]) => userId),
  );

  const photosByUserId = photos.reduce((groups, photo) => {
    if (!photo.user_id || (!photo.file_id && !photo.photo_url)) return groups;
    groups[photo.user_id] = groups[photo.user_id] || [];
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
      const primaryPhoto =
        userPhotos.find((photo) => String(photo.is_primary).toUpperCase() === "TRUE") ||
        userPhotos[0] ||
        {};

      return {
        id: user.user_id,
        name: user.nickname || user.user_id,
        nickname: user.nickname || user.user_id,
        email: user.email,
        gender: user.gender,
        age: user.age,
        birthYear: user.birth_year || (user.age ? String(new Date().getFullYear() - Number(user.age)) : ""),
        birthMonth: user.birth_month || "",
        birthDay: user.birth_day || "",
        city: user.city,
        district: user.district,
        area: [user.city, user.district].filter(Boolean).join("・"),
        looking: user.coffee_goal,
        coffeeGoal: user.coffee_goal,
        occupation: user.occupation,
        education: user.education,
        relationshipStatus: user.relationship_status,
        hasChildren: user.has_children,
        smoking: user.smoking,
        drinking: user.drinking,
        intro: user.intro,
        time: user.available_times,
        availabilityNote: user.available_times,
        place: place.place_name || "",
        meetingArea: place.place_name || "",
        meetingPlaceName: place.place_name || "",
        meetingPlaceUrl: place.map_url || "",
        meetingLat: place.lat || "",
        meetingLng: place.lng || "",
        interestKeywords: user.interest_keywords,
      photo: driveImageUrl(primaryPhoto.file_id, primaryPhoto.photo_url),
      photos: userPhotos.map((photo) => driveImageUrl(photo.file_id, photo.photo_url)).filter(Boolean),
      };
    });

  return {
    ok: true,
    candidates,
    invites: invites.map(function(invite, index) {
      return {
        id: invite.invite_id || "sheet-invite-" + (index + 1),
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
          invite.accepted_at ? "已確認：" + invite.accepted_at : "",
          invite.selected_times ? "可選時段：" + invite.selected_times : "",
          invite.place_name ? "地點：" + invite.place_name : "",
        ].filter(Boolean).join("；"),
      };
    }),
  };
}

function saveInvite(invite) {
  var senderUserId = String(invite.senderUserId || invite.sender_user_id || "").trim();
  var receiverUserId = String(invite.receiverUserId || invite.receiver_user_id || "").trim();
  if (!senderUserId || !receiverUserId) return { ok: false, error: "senderUserId and receiverUserId are required" };

  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("invites");
  if (!sheet) return { ok: false, error: "invites sheet not found" };
  var headers = ensureInviteHeaders(sheet);
  var today = new Date().toISOString().slice(0, 10);
  var note = stripAddressLines(invite.note).trim();
  var selectedTimes = Array.isArray(invite.selectedTimes)
    ? invite.selectedTimes.join("、")
    : String(invite.selected_times || invite.selectedTimes || detailFromNote(note, "可選時段") || "").trim();
  var placeName = String(invite.placeName || invite.place_name || detailFromNote(note, "地點") || "").trim();
  var feedbackScore = String(invite.feedbackScore || invite.feedback_score || detailFromNote(note, "問卷分數").replace(/分$/, "") || "").trim();
  var feedbackGender = String(invite.feedbackGender || invite.feedback_gender || "").trim();
  var maleFeedbackScore = String(invite.maleFeedbackScore || invite.male_feedback_score || (feedbackGender === "男性" ? feedbackScore : "")).trim();
  var femaleFeedbackScore = String(invite.femaleFeedbackScore || invite.female_feedback_score || (feedbackGender === "女性" ? feedbackScore : "")).trim();
  var hasParticipantFeedbackScore = Boolean(maleFeedbackScore || femaleFeedbackScore);
  var suppliedInviteId = String(invite.inviteId || invite.invite_id || invite.id || "").trim();
  var inviteId = suppliedInviteId || "invite-" + Date.now().toString(36);
  var rows = sheetRows("invites");
  var requestedStatus = statusToInviteStatus(invite.status);
  var existingIndex = suppliedInviteId
    ? rows.findIndex(function(row) {
        return String(row.invite_id || "").trim() === suppliedInviteId;
      })
    : -1;
  if (existingIndex < 0 && !suppliedInviteId) {
    existingIndex = rows.findIndex(function(row) {
      var status = statusToInviteStatus(row.status);
      return row.sender_user_id === senderUserId &&
        row.receiver_user_id === receiverUserId &&
        ["sent", "pending", "incoming", "confirmed"].indexOf(status) >= 0;
    });
  }
  if (existingIndex >= 0 && ["confirmed", "cancelled", "done"].indexOf(requestedStatus) >= 0) {
    var existing = rows[existingIndex];
    var acceptedTime = String(invite.acceptedTime || invite.accepted_at || detailFromNote(note, "已確認") || "").trim();
    var updated = Object.assign({}, existing, {
      status: requestedStatus,
      note: note || stripAddressLines(requestedStatus === "confirmed" ? "已確認：" + acceptedTime + "；" + (existing.note || "") : existing.note || ""),
      feedback_score: hasParticipantFeedbackScore ? existing.feedback_score || "" : feedbackScore || existing.feedback_score || "",
      male_feedback_score: maleFeedbackScore || existing.male_feedback_score || "",
      female_feedback_score: femaleFeedbackScore || existing.female_feedback_score || "",
      accepted_at: acceptedTime || existing.accepted_at || "",
      updated_at: today,
    });
    sheet.getRange(existingIndex + 2, 1, 1, headers.length).setValues([headers.map(function(header) {
      return updated[header] !== undefined ? updated[header] : "";
    })]);
    SpreadsheetApp.flush();
    if (requestedStatus === "confirmed" && statusToInviteStatus(existing.status) !== "confirmed") {
      sendInviteConfirmedEmails(updated);
    }
    return {
      ok: true,
      inviteId: rows[existingIndex].invite_id || inviteId,
      updated: true,
      rowIndex: existingIndex + 1,
      beforeStatus: existing.status || "",
      afterStatus: requestedStatus,
    };
  }
  if (existingIndex < 0 && ["confirmed", "cancelled", "done"].indexOf(requestedStatus) >= 0) {
    return {
      ok: false,
      error: "Invite update target not found",
      inviteId: inviteId,
      requestedStatus: requestedStatus,
      suppliedInviteId: suppliedInviteId,
      matchingIds: rows.map(function(row) { return row.invite_id || ""; }),
    };
  }
  if (existingIndex >= 0) {
    return { ok: true, inviteId: rows[existingIndex].invite_id || inviteId, duplicate: true };
  }

  var values = {
    invite_id: inviteId,
    sender_user_id: senderUserId,
    receiver_user_id: receiverUserId,
    status: "sent",
    selected_times: selectedTimes,
    place_name: placeName,
    note: note || "可選時段：" + selectedTimes + "；地點：" + placeName,
    feedback_score: hasParticipantFeedbackScore ? "" : feedbackScore,
    male_feedback_score: maleFeedbackScore,
    female_feedback_score: femaleFeedbackScore,
    created_at: today,
    updated_at: today,
    accepted_at: "",
  };
  sheet.appendRow(headers.map(function(header) {
    return values[header] !== undefined ? values[header] : "";
  }));
  SpreadsheetApp.flush();
  sendInviteCreatedEmail(values);
  return { ok: true, inviteId: inviteId };
}

function sendInviteCreatedEmail(invite) {
  try {
    var usersById = usersByUserId();
    var sender = usersById[invite.sender_user_id] || {};
    var receiver = usersById[invite.receiver_user_id] || {};
    if (!receiver.email) return;

    var senderName = displayUserName(sender, "對方");
    var viewUrl = inviteViewUrl(invite.invite_id, receiver.email);
    var htmlBody = emailLayout(
      "有人邀請你喝咖啡 ☕",
      [
        detailRow("對方姓名", senderName),
        detailRow("咖啡店", invite.place_name || "尚未填寫"),
        detailRow("對方提供的可約時間", invite.selected_times || "尚未填寫"),
      ].join(""),
      emailButton("查看邀約", viewUrl)
    );

    sendCoffeeMeetEmail({
      to: receiver.email,
      subject: "有人邀請你喝咖啡 ☕",
      htmlBody: htmlBody,
      body: senderName + " 邀請你喝咖啡。\n咖啡店：" + (invite.place_name || "尚未填寫") + "\n可約時間：" + (invite.selected_times || "尚未填寫") + "\n查看邀約：" + viewUrl,
    });
  } catch (error) {
    console.warn("sendInviteCreatedEmail failed", error);
  }
}

function sendInviteConfirmedEmails(invite) {
  try {
    var usersById = usersByUserId();
    var sender = usersById[invite.sender_user_id] || {};
    var receiver = usersById[invite.receiver_user_id] || {};
    sendInviteConfirmedEmailTo(sender, receiver, invite);
    sendInviteConfirmedEmailTo(receiver, sender, invite);
  } catch (error) {
    console.warn("sendInviteConfirmedEmails failed", error);
  }
}

function sendInviteConfirmedEmailTo(recipient, counterpart, invite) {
  if (!recipient.email) return;
  var counterpartName = displayUserName(counterpart, "對方");
  var viewUrl = inviteViewUrl(invite.invite_id, recipient.email);
  var notes = [
    "第一次請約公開咖啡店",
    "若需取消請提早通知",
    "見面完成後請記得填寫回饋",
  ];
  var htmlBody = emailLayout(
    "咖啡邀約已成立 ☕",
    [
      detailRow("對方姓名", counterpartName),
      detailRow("見面地點", invite.place_name || "尚未填寫"),
      detailRow("見面時間", invite.accepted_at || invite.selected_times || "尚未填寫"),
      '<div style="margin-top:18px;font-weight:700;color:#312820;">注意事項</div>',
      '<ul style="margin:8px 0 0 20px;padding:0;color:#4f463d;line-height:1.8;">' +
        notes.map(function(note) { return "<li>" + escapeHtml(note) + "</li>"; }).join("") +
      "</ul>",
    ].join(""),
    emailButton("查看邀約", viewUrl)
  );

  sendCoffeeMeetEmail({
    to: recipient.email,
    subject: "咖啡邀約已成立 ☕",
    htmlBody: htmlBody,
    body: "你和 " + counterpartName + " 的咖啡邀約已成立。\n見面地點：" + (invite.place_name || "尚未填寫") + "\n見面時間：" + (invite.accepted_at || invite.selected_times || "尚未填寫") + "\n注意事項：\n- " + notes.join("\n- ") + "\n查看邀約：" + viewUrl,
  });
}

function sendCoffeeMeetEmail(message) {
  try {
    MailApp.sendEmail(message);
  } catch (error) {
    console.warn("MailApp.sendEmail failed; sending fallback email", error);
    MailApp.sendEmail({
      to: EMAIL_FALLBACK_TO,
      subject: "[CoffeeMeet 寄信失敗] " + (message.subject || ""),
      htmlBody: '<div style="font-family:Arial,Helvetica,sans-serif;line-height:1.6;">' +
        '<p><strong>原收件者：</strong>' + escapeHtml(message.to || "") + '</p>' +
        '<p><strong>錯誤：</strong>' + escapeHtml(error && error.message ? error.message : String(error)) + '</p>' +
        '<hr />' +
        (message.htmlBody || "").replace(/<script[\s\S]*?<\/script>/gi, "") +
      '</div>',
      body: "原收件者：" + (message.to || "") +
        "\n錯誤：" + (error && error.message ? error.message : String(error)) +
        "\n\n--- 原信內容 ---\n" + (message.body || ""),
    });
  }
}

function usersByUserId() {
  return sheetRows("users").reduce(function(map, user) {
    if (user.user_id) map[user.user_id] = user;
    return map;
  }, {});
}

function displayUserName(user, fallback) {
  return user.nickname || user.name || user.email || fallback;
}

function inviteViewUrl(inviteId, email) {
  var query = [
    "inviteId=" + encodeURIComponent(inviteId || ""),
    email ? "email=" + encodeURIComponent(email) : "",
  ].filter(Boolean).join("&");
  return APP_URL + (query ? "?" + query : "");
}

function emailLayout(title, contentHtml, actionHtml) {
  return '<div style="margin:0;padding:28px;background:#f7f2ec;font-family:Arial,Helvetica,sans-serif;color:#312820;">' +
    '<div style="max-width:560px;margin:0 auto;background:#fff;border:1px solid #eadfd3;border-radius:8px;padding:28px;">' +
      '<h1 style="margin:0 0 20px;font-size:24px;line-height:1.35;color:#2f241b;">' + escapeHtml(title) + '</h1>' +
      '<div style="font-size:15px;line-height:1.7;">' + contentHtml + '</div>' +
      '<div style="margin-top:24px;">' + actionHtml + '</div>' +
    '</div>' +
  '</div>';
}

function detailRow(label, value) {
  return '<div style="margin:0 0 12px;">' +
    '<div style="font-size:13px;color:#8a7562;">' + escapeHtml(label) + '</div>' +
    '<div style="font-size:16px;font-weight:700;color:#312820;">' + escapeHtml(value || "尚未填寫") + '</div>' +
  '</div>';
}

function emailButton(label, url) {
  return '<a href="' + escapeHtml(url) + '" style="display:inline-block;background:#5b3f2f;color:#fff;text-decoration:none;border-radius:6px;padding:12px 18px;font-weight:700;">' + escapeHtml(label) + '</a>';
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function statusToInviteStatus(status) {
  var normalized = String(status || "").trim().toLowerCase();
  if (["incoming", "sent", "confirmed", "done", "cancelled"].indexOf(normalized) >= 0) return normalized;
  if (["accepted", "scheduled", "ready", "\u5df2\u78ba\u8a8d", "\u5f85\u898b\u9762"].indexOf(normalized) >= 0) return "confirmed";
  if (["completed", "complete", "\u5df2\u5b8c\u6210"].indexOf(normalized) >= 0) return "done";
  if (["canceled", "cancelled", "\u5df2\u53d6\u6d88"].indexOf(normalized) >= 0) return "cancelled";
  if (["pending", "sent", "\u5df2\u9080\u7d04"].indexOf(normalized) >= 0) return "sent";
  if (["\u88ab\u9080\u7d04"].indexOf(normalized) >= 0) return "incoming";
  return normalized || "sent";
}

function detailFromNote(note, label) {
  var match = String(note || "").match(new RegExp(label + "：([^；]+)"));
  return match ? match[1].trim() : "";
}

function saveUserProfile(profile) {
  const email = String(profile.email || "").trim().toLowerCase();
  if (!email) return { ok: false, error: "Email is required" };

  const usersSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("users");
  const headers = ensureUserHeaders(usersSheet);
  const users = sheetRows("users");
  const existingIndex = users.findIndex((user) => String(user.email || "").toLowerCase() === email);
  const existing = existingIndex >= 0 ? users[existingIndex] : null;
  const userId = existing
    ? existing.user_id
    : `user-${email.replace(/@.*/, "").replace(/[^a-z0-9]+/gi, "-")}-${Date.now().toString(36)}`;
  const placeId = existing && existing.meeting_place_id
    ? existing.meeting_place_id
    : `place-${email.replace(/@.*/, "").replace(/[^a-z0-9]+/gi, "-")}`;
  const today = new Date().toISOString().slice(0, 10);
  const age = calculateAge(profile.birthYear, profile.birthMonth, profile.birthDay);

  const values = {
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
    available_times: profile.availabilityNote || "",
    meeting_place_id: placeId,
    lock_until: existing ? existing.lock_until || "" : "",
    status: "active",
    created_at: existing ? existing.created_at || today : today,
    updated_at: today,
    notes: existing ? existing.notes || "" : "",
  };

  const row = headers.map((header) => (values[header] !== undefined ? values[header] : ""));
  if (existing) {
    usersSheet.getRange(existingIndex + 2, 1, 1, headers.length).setValues([row]);
  } else {
    usersSheet.appendRow(row);
  }

  saveMeetingPlace({
    placeId,
    userId,
    profile,
    today,
  });

  let photoCount = 0;
  let photoError = "";
  try {
    photoCount = saveUserPhotos({
      userId,
      profile,
      today,
    });
  } catch (error) {
    photoError = error.message || String(error);
  }

  return { ok: true, created: !existing, userId, photoCount, photoError };
}

function calculateAge(birthYear, birthMonth, birthDay) {
  const year = Number(birthYear);
  if (!year) return "";
  const today = new Date();
  let age = today.getFullYear() - year;
  const month = Number(birthMonth);
  const day = Number(birthDay);
  if (month && day) {
    const birthdayThisYear = new Date(today.getFullYear(), month - 1, day);
    if (today < birthdayThisYear) age -= 1;
  }
  return String(age);
}

function saveUserPhotos({ userId, profile, today }) {
  const photos = Array.isArray(profile.photos) ? profile.photos.filter(Boolean).slice(0, 3) : [];
  const photosSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("user_photos");
  if (!photosSheet) return 0;
  const headers = ensurePhotoHeaders(photosSheet);
  const rows = sheetRows("user_photos");
  const existingIndexes = rows
    .map((photo, index) => ({ photo, index }))
    .filter(({ photo }) => photo.user_id === userId);

  photos.forEach((photoValue, index) => {
    const existing = existingIndexes[index] && existingIndexes[index].photo;
    const existingRowIndex = existingIndexes[index] && existingIndexes[index].index;
    const uploaded = savePhotoFile(photoValue, {
      userId,
      order: index + 1,
      gender: profile.gender,
    });
    const values = {
      photo_id: existing && existing.photo_id ? existing.photo_id : `photo-${userId}-${index + 1}`,
      user_id: userId,
      photo_url: "",
      file_id: uploaded.fileId || (existing && existing.file_id) || "",
      photo_order: String(index + 1),
      is_primary: index === 0 ? "TRUE" : "FALSE",
      status: "active",
      created_at: existing && existing.created_at ? existing.created_at : today,
      updated_at: today,
    };
    const row = headers.map((header) => (values[header] !== undefined ? values[header] : ""));
    if (existingRowIndex !== undefined) {
      photosSheet.getRange(existingRowIndex + 2, 1, 1, headers.length).setValues([row]);
    } else {
      photosSheet.appendRow(row);
    }
  });

  for (let index = photos.length; index < existingIndexes.length; index += 1) {
    const existing = existingIndexes[index].photo;
    const rowIndex = existingIndexes[index].index;
    const values = Object.assign({}, existing, {
      status: "deleted",
      updated_at: today,
    });
    const row = headers.map((header) => (values[header] !== undefined ? values[header] : ""));
    photosSheet.getRange(rowIndex + 2, 1, 1, headers.length).setValues([row]);
  }

  return photos.length;
}

function savePhotoFile(photoValue, { userId, order, gender }) {
  const dataUrlMatch = String(photoValue || "").match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!dataUrlMatch) return { photoUrl: "", fileId: extractDriveFileId(photoValue) };

  const mimeType = dataUrlMatch[1];
  const extension = mimeType.split("/")[1].replace("jpeg", "jpg");
  const bytes = Utilities.base64Decode(dataUrlMatch[2]);
  const fileName = `${userId}-${order}-${Date.now()}.${extension}`;
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const folderId = String(gender || "").includes("女") ? GIRL_FOLDER_ID : BOY_FOLDER_ID;
  const file = DriveApp.getFolderById(folderId).createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {
    photoUrl: "",
    fileId: file.getId(),
  };
}

function getVersion() {
  return {
    ok: true,
    version: "profile-persistence-2026-06-24-5",
    features: ["birthday-age", "photo-drive-upload", "photo-error-isolated", "location-owner-fallback", "photo-delete-all", "photo-storage-diagnostics", "drive-authorization-helper"],
  };
}

function authorizeDriveAccess() {
  const spreadsheetName = SpreadsheetApp.openById(SPREADSHEET_ID).getName();
  const boyFolderName = DriveApp.getFolderById(BOY_FOLDER_ID).getName();
  const girlFolderName = DriveApp.getFolderById(GIRL_FOLDER_ID).getName();
  return {
    ok: true,
    spreadsheetName,
    boyFolderName,
    girlFolderName,
  };
}

function getPhotoStorageMeta() {
  return {
    ok: true,
    boy: folderMeta(BOY_FOLDER_ID),
    girl: folderMeta(GIRL_FOLDER_ID),
  };
}

function publishSeedPhotos(token) {
  const expectedToken = PropertiesService.getScriptProperties().getProperty("SEED_MIGRATION_TOKEN");
  if (!expectedToken || token !== expectedToken) return { ok: false, error: "Invalid migration token" };
  return {
    ok: true,
    girl: publishFolderPhotos(GIRL_FOLDER_ID),
    boy: publishFolderPhotos(BOY_FOLDER_ID),
  };
}

function publishFolderPhotos(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const published = [];
  while (files.hasNext()) {
    const file = files.next();
    if (!String(file.getMimeType() || "").startsWith("image/")) continue;
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    published.push({
      id: file.getId(),
      name: file.getName(),
      url: driveImageUrl(file.getId()),
    });
  }
  return published;
}

function clearPhotoUrls(token) {
  const expectedToken = PropertiesService.getScriptProperties().getProperty("SEED_MIGRATION_TOKEN");
  if (!expectedToken || token !== expectedToken) return { ok: false, error: "Invalid migration token" };

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("user_photos");
  if (!sheet) return { ok: false, error: "user_photos sheet not found" };
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || [];
  const photoUrlIndex = headers.indexOf("photo_url");
  const fileIdIndex = headers.indexOf("file_id");
  if (photoUrlIndex < 0 || fileIdIndex < 0) return { ok: false, error: "photo_url or file_id column not found" };

  let cleared = 0;
  for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    const row = values[rowIndex];
    if (!row[fileIdIndex] || !row[photoUrlIndex]) continue;
    sheet.getRange(rowIndex + 1, photoUrlIndex + 1).setValue("");
    cleared += 1;
  }
  return { ok: true, cleared };
}

function folderMeta(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    let fileCount = 0;
    const files = folder.getFiles();
    while (files.hasNext() && fileCount < 5) {
      files.next();
      fileCount += 1;
    }
    return {
      ok: true,
      id: folderId,
      name: folder.getName(),
      url: folder.getUrl(),
      hasFiles: fileCount > 0,
      sampledFileCount: fileCount,
    };
  } catch (error) {
    return {
      ok: false,
      id: folderId,
      error: error.message || String(error),
    };
  }
}

function extractDriveFileId(value) {
  const text = String(value || "").trim();
  const ucMatch = text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (ucMatch) return ucMatch[1];
  const fileMatch = text.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (fileMatch) return fileMatch[1];
  return /^[a-zA-Z0-9_-]{20,}$/.test(text) ? text : "";
}

function saveMeetingPlace({ placeId, userId, profile, today }) {
  const placesSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("meeting_places");
  if (!placesSheet) return;
  const headers = ensureMeetingPlaceHeaders(placesSheet);
  const places = sheetRows("meeting_places");
  const existingIndex = places.findIndex((place) => place.place_id === placeId || place.owner_user_id === userId);
  const existing = existingIndex >= 0 ? places[existingIndex] : null;
  const coords = firstLineCoords(profile.meetingArea);
  const sanitizedMeetingArea = stripAddressLines(profile.meetingArea);
  const values = {
    place_id: placeId,
    place_name: profile.meetingPlaceName || firstLineValue(sanitizedMeetingArea, "地點") || sanitizedMeetingArea || "",
    city: profile.city || "",
    district: profile.district || "",
    map_url: profile.meetingPlaceUrl || firstLineValue(profile.meetingArea, "Google Maps") || "",
    lat: profile.meetingLat || (coords && coords.lat) || "",
    lng: profile.meetingLng || (coords && coords.lng) || "",
    google_place_id: existing ? existing.google_place_id || "" : "",
    category: existing ? existing.category || "coffee" : "coffee",
    is_public: existing ? existing.is_public || "FALSE" : "FALSE",
    owner_user_id: userId,
    created_at: existing ? existing.created_at || today : today,
    updated_at: today,
    raw_input: sanitizedMeetingArea,
  };
  const row = headers.map((header) => (values[header] !== undefined ? values[header] : ""));
  if (existing) {
    placesSheet.getRange(existingIndex + 2, 1, 1, headers.length).setValues([row]);
  } else {
    placesSheet.appendRow(row);
  }
}

function firstLineValue(value, label) {
  const match = String(value || "").match(new RegExp("^" + label + "：(.+)$", "m"));
  return match ? match[1].trim() : "";
}

function stripAddressLines(value) {
  return String(value || "")
    .split(/\r?\n/)
    .filter(function(line) {
      return !/^地址：/.test(line.trim());
    })
    .join("\n");
}

function firstLineCoords(value) {
  const match = String(value || "").match(/^座標：\s*(-?\d+(?:\.\d+)?),\s*(-?\d+(?:\.\d+)?)/m);
  return match ? { lat: match[1], lng: match[2] } : null;
}

function ensureMeetingPlaceHeaders(sheet) {
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
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].filter(Boolean);
  if (!headers.length) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return requiredHeaders;
  }
  const missingHeaders = requiredHeaders.filter((header) => !headers.includes(header));
  if (missingHeaders.length) {
    sheet.getRange(1, headers.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
  }
  return headers.concat(missingHeaders);
}

function ensurePhotoHeaders(sheet) {
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
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].filter(Boolean);
  if (!headers.length) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return requiredHeaders;
  }

  const missingHeaders = requiredHeaders.filter((header) => !headers.includes(header));
  if (missingHeaders.length) {
    sheet.getRange(1, headers.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
  }
  return headers.concat(missingHeaders);
}

function ensureUserHeaders(sheet) {
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
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].filter(Boolean);
  if (!headers.length) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return requiredHeaders;
  }

  const missingHeaders = requiredHeaders.filter((header) => !headers.includes(header));
  if (missingHeaders.length) {
    sheet.getRange(1, headers.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
  }
  return headers.concat(missingHeaders);
}

function ensureInviteHeaders(sheet) {
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
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].filter(Boolean);
  if (!headers.length) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return requiredHeaders;
  }

  const missingHeaders = requiredHeaders.filter((header) => !headers.includes(header));
  if (missingHeaders.length) {
    sheet.getRange(1, headers.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
  }
  return headers.concat(missingHeaders);
}

function driveImageUrl(fileId, fallbackUrl) {
  if (fileId) return `https://lh3.googleusercontent.com/d/${fileId}=w1200`;
  return fallbackUrl || "";
}
