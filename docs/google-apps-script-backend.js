const SPREADSHEET_ID = "1JzQwlNWQphrHfUCcpQl3pC6jHLXN6ykwvJKSs5Uaxsg";
const BOY_FOLDER_ID = "12M91XLYJ7seBNhrsGL2JZNJZlSGlL9ri";
const GIRL_FOLDER_ID = "1r_6bPvf-R5lHM92g379al71G8zzmj-of";

function doGet(event) {
  const action = event.parameter.action || "";
  if (action === "app-data") return jsonResponse(getAppData());
  if (action === "sheets") return jsonResponse(getSheetMeta());
  if (action === "version") return jsonResponse(getVersion());
  if (action === "photo-storage") return jsonResponse(getPhotoStorageMeta());
  return jsonResponse({ ok: false, error: "Unknown action" });
}

function doPost(event) {
  try {
    const action = event.parameter.action || "";
    const body = JSON.parse(event.postData.contents || "{}");
    if (action === "users") return jsonResponse(saveUserProfile(body));
    return jsonResponse({ ok: false, error: "Unknown action" });
  } catch (error) {
    return jsonResponse({ ok: false, error: error.message || String(error) });
  }
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
  const placesByOwnerUserId = Object.fromEntries(
    places.map((place) => [place.owner_user_id, place]).filter(([userId]) => userId),
  );

  const photosByUserId = photos.reduce((groups, photo) => {
    if (!photo.user_id || !photo.photo_url) return groups;
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
      photo_url: uploaded.photoUrl || photoValue,
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
  if (!dataUrlMatch) return { photoUrl: photoValue, fileId: extractDriveFileId(photoValue) };

  const mimeType = dataUrlMatch[1];
  const extension = mimeType.split("/")[1].replace("jpeg", "jpg");
  const bytes = Utilities.base64Decode(dataUrlMatch[2]);
  const fileName = `${userId}-${order}-${Date.now()}.${extension}`;
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const folderId = String(gender || "").includes("女") ? GIRL_FOLDER_ID : BOY_FOLDER_ID;
  const file = DriveApp.getFolderById(folderId).createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {
    photoUrl: driveImageUrl(file.getId()),
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
  const values = {
    place_id: placeId,
    place_name: profile.meetingPlaceName || firstLineValue(profile.meetingArea, "地點") || profile.meetingArea || "",
    address: existing ? existing.address || "" : "",
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
    raw_input: profile.meetingArea || "",
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

function firstLineCoords(value) {
  const match = String(value || "").match(/^座標：\s*(-?\d+(?:\.\d+)?),\s*(-?\d+(?:\.\d+)?)/m);
  return match ? { lat: match[1], lng: match[2] } : null;
}

function ensureMeetingPlaceHeaders(sheet) {
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

function driveImageUrl(fileId, fallbackUrl) {
  if (fileId) return `https://drive.google.com/thumbnail?id=${fileId}&sz=w1200`;
  return fallbackUrl || "";
}
