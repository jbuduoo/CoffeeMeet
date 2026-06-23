const SPREADSHEET_ID = "1JzQwlNWQphrHfUCcpQl3pC6jHLXN6ykwvJKSs5Uaxsg";
const BOY_FOLDER_ID = "12M91XLYJ7seBNhrsGL2JZNJZlSGlL9ri";
const GIRL_FOLDER_ID = "1r_6bPvf-R5lHM92g379al71G8zzmj-of";

function doGet(event) {
  const action = event.parameter.action || "";
  if (action === "app-data") return jsonResponse(getAppData());
  if (action === "sheets") return jsonResponse(getSheetMeta());
  return jsonResponse({ ok: false, error: "Unknown action" });
}

function doPost(event) {
  const action = event.parameter.action || "";
  const body = JSON.parse(event.postData.contents || "{}");
  if (action === "users") return jsonResponse(saveUserProfile(body));
  return jsonResponse({ ok: false, error: "Unknown action" });
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
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

  const photosByUserId = photos.reduce((groups, photo) => {
    if (!photo.user_id || !photo.photo_url) return groups;
    groups[photo.user_id] = groups[photo.user_id] || [];
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
        interestKeywords: user.interest_keywords,
        photo: driveImageUrl(primaryPhoto.file_id, primaryPhoto.photo_url),
        photos: userPhotos.map((photo) => driveImageUrl(photo.file_id, photo.photo_url)).filter(Boolean),
      };
    });

  return { ok: true, candidates, invites };
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
  const today = new Date().toISOString().slice(0, 10);

  const values = {
    user_id: userId,
    nickname: profile.nickname || email,
    gender: profile.gender || "",
    age: profile.birthYear ? String(new Date().getFullYear() - Number(profile.birthYear)) : "",
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
    meeting_place_id: existing ? existing.meeting_place_id || "" : "",
    lock_until: existing ? existing.lock_until || "" : "",
    status: "active",
    created_at: existing ? existing.created_at || today : today,
    updated_at: today,
    notes: existing ? existing.notes || "" : "",
  };

  const row = headers.map((header) => (values[header] !== undefined ? values[header] : ""));
  if (existing) {
    usersSheet.getRange(existingIndex + 2, 1, 1, headers.length).setValues([row]);
    return { ok: true, created: false, userId };
  }

  usersSheet.appendRow(row);
  return { ok: true, created: true, userId };
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

function driveImageUrl(fileId, fallbackUrl) {
  if (fileId) return `https://drive.google.com/uc?export=view&id=${fileId}`;
  return fallbackUrl || "";
}
