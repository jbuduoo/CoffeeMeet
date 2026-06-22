const fs = require("fs");
const crypto = require("crypto");

const defaultCredentialPath = "coffeemeet-500009-05662f9fdf8a.json";
const defaultSpreadsheetId = "1JzQwlNWQphrHfUCcpQl3pC6jHLXN6ykwvJKSs5Uaxsg";

function base64url(value) {
  return Buffer.from(value)
    .toString("base64")
    .replace(/=/g, "")
    .replace(/\+/g, "-")
    .replace(/\//g, "_");
}

function signJwt(credential) {
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const claim = {
    iss: credential.client_email,
    scope: "https://www.googleapis.com/auth/spreadsheets.readonly",
    aud: credential.token_uri,
    exp: now + 3600,
    iat: now,
  };

  const unsigned = `${base64url(JSON.stringify(header))}.${base64url(JSON.stringify(claim))}`;
  const signature = crypto.createSign("RSA-SHA256").update(unsigned).sign(credential.private_key);
  return `${unsigned}.${signature
    .toString("base64")
    .replace(/=/g, "")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")}`;
}

function loadCredential(credentialPath = defaultCredentialPath) {
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    return JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  }

  if (!fs.existsSync(credentialPath)) {
    throw new Error(`找不到憑證檔：${credentialPath}`);
  }
  return JSON.parse(fs.readFileSync(credentialPath, "utf8"));
}

async function getAccessToken(credential) {
  const response = await fetch(credential.token_uri, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: signJwt(credential),
    }),
  });

  const text = await response.text();
  if (!response.ok) {
    throw new Error(`Token error ${response.status}: ${text}`);
  }
  return JSON.parse(text).access_token;
}

async function requestSheetsApi(path, accessToken) {
  const response = await fetch(`https://sheets.googleapis.com/v4/${path}`, {
    headers: { authorization: `Bearer ${accessToken}` },
  });

  const text = await response.text();
  if (!response.ok) {
    throw new Error(`Sheets error ${response.status}: ${text}`);
  }
  return JSON.parse(text);
}

async function getSpreadsheetMeta(options = {}) {
  const credential = loadCredential(options.credentialPath);
  const accessToken = await getAccessToken(credential);
  const spreadsheetId = options.spreadsheetId || defaultSpreadsheetId;
  const fields = "spreadsheetId,properties.title,sheets.properties.title";
  const spreadsheet = await requestSheetsApi(`spreadsheets/${spreadsheetId}?fields=${fields}`, accessToken);

  return {
    serviceAccount: credential.client_email,
    spreadsheetId: spreadsheet.spreadsheetId,
    title: spreadsheet.properties && spreadsheet.properties.title,
    sheets: (spreadsheet.sheets || []).map((sheet) => sheet.properties.title),
  };
}

async function getSheetValues(range, options = {}) {
  const credential = loadCredential(options.credentialPath);
  const accessToken = await getAccessToken(credential);
  const spreadsheetId = options.spreadsheetId || defaultSpreadsheetId;
  return requestSheetsApi(
    `spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`,
    accessToken,
  );
}

module.exports = {
  defaultCredentialPath,
  defaultSpreadsheetId,
  getSpreadsheetMeta,
  getSheetValues,
};
