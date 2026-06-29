const {
  batchUpdateSpreadsheet,
  clearSheetValues,
  getSheetValues,
  getSpreadsheetMeta,
  updateSheetValues,
} = require("../sheets-client");

const LOG_SHEET = "上來紀錄";
const STATS_SHEET = "人員統計";

function rowsToObjects(values = []) {
  const [headers = [], ...rows] = values;
  return rows
    .filter((row) => row.some((cell) => String(cell || "").trim()))
    .map((row) =>
      Object.fromEntries(headers.map((header, index) => [String(header || "").trim(), row[index] || ""])),
    );
}

function quoteSheetName(sheetName) {
  return `'${String(sheetName).replace(/'/g, "''")}'`;
}

async function ensureSheets() {
  const meta = await getSpreadsheetMeta();
  const existing = new Set(meta.sheets);
  const requests = [LOG_SHEET, STATS_SHEET]
    .filter((title) => !existing.has(title))
    .map((title) => ({
      addSheet: {
        properties: {
          title,
          gridProperties: {
            rowCount: title === LOG_SHEET ? 1000 : 500,
            columnCount: title === LOG_SHEET ? 6 : 13,
            frozenRowCount: 1,
          },
        },
      },
    }));

  if (requests.length) {
    await batchUpdateSpreadsheet(requests);
  }

  return getSpreadsheetMeta();
}

async function formatSheets(meta) {
  const sheetByTitle = Object.fromEntries((meta.sheetProperties || []).map((sheet) => [sheet.title, sheet]));
  const requests = [LOG_SHEET, STATS_SHEET]
    .map((title) => sheetByTitle[title])
    .filter(Boolean)
    .flatMap((sheet) => [
      {
        repeatCell: {
          range: {
            sheetId: sheet.sheetId,
            startRowIndex: 0,
            endRowIndex: 1,
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 0.9, green: 0.95, blue: 1 },
              textFormat: { bold: true },
              horizontalAlignment: "CENTER",
            },
          },
          fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
        },
      },
      {
        autoResizeDimensions: {
          dimensions: {
            sheetId: sheet.sheetId,
            dimension: "COLUMNS",
            startIndex: 0,
            endIndex: sheet.title === LOG_SHEET ? 6 : 13,
          },
        },
      },
    ]);

  if (requests.length) {
    await batchUpdateSpreadsheet(requests);
  }
}

function buildStatsRows(users) {
  const log = quoteSheetName(LOG_SHEET);
  const headers = [
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

  const sortedUsers = [...users].sort((a, b) =>
    String(a.nickname || a.email || a.user_id).localeCompare(String(b.nickname || b.email || b.user_id), "zh-Hant"),
  );

  const rows = sortedUsers.map((user, index) => {
    const rowNumber = index + 2;
    return [
      user.user_id || "",
      user.nickname || user.email || user.user_id || "",
      user.email || "",
      user.status || "",
      user.created_at || "",
      user.updated_at || "",
      `=COUNTIF(${log}!B:B,A${rowNumber})`,
      `=IF(G${rowNumber}=0,"",MAXIFS(${log}!A:A,${log}!B:B,A${rowNumber}))`,
      `=IF(G${rowNumber}=0,"",MINIFS(${log}!A:A,${log}!B:B,A${rowNumber}))`,
      `=COUNTIFS(${log}!B:B,A${rowNumber},${log}!A:A,">="&TODAY()-7)`,
      `=COUNTIFS(${log}!B:B,A${rowNumber},${log}!A:A,">="&TODAY()-30)`,
      `=IF(G${rowNumber}=0,"未記錄",IF(H${rowNumber}>=TODAY()-7,"活躍",IF(H${rowNumber}>=TODAY()-30,"普通","沉寂")))`,
      "",
    ];
  });

  return [headers, ...rows];
}

async function main() {
  const meta = await ensureSheets();
  const usersSheet = await getSheetValues("users!A1:AZ2000");
  const users = rowsToObjects(usersSheet.values);

  await updateSheetValues(`${LOG_SHEET}!A1:F1`, [[
    "上來時間",
    "user_id",
    "姓名/暱稱",
    "email",
    "來源/活動",
    "備註",
  ]]);

  await clearSheetValues(`${STATS_SHEET}!A:M`);
  await updateSheetValues(`${STATS_SHEET}!A1:M${Math.max(users.length + 1, 1)}`, buildStatsRows(users));
  await formatSheets(meta);

  console.log(`已更新 ${STATS_SHEET}：${users.length} 位使用者`);
  console.log(`${LOG_SHEET} 可手動新增每次上來紀錄；${STATS_SHEET} 會自動統計。`);
}

main().catch((error) => {
  console.error(error.message || error);
  process.exit(1);
});
