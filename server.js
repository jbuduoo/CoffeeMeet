const http = require("http");
const fs = require("fs");
const path = require("path");
const { getSpreadsheetMeta, getSheetValues } = require("./sheets-client");

const port = Number(process.env.PORT || 3000);
const rootDir = __dirname;

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
