const { getSpreadsheetMeta } = require("./sheets-client");

async function main() {
  const spreadsheet = await getSpreadsheetMeta();

  console.log("連線成功");
  console.log(`Service account: ${spreadsheet.serviceAccount}`);
  console.log(`Spreadsheet: ${spreadsheet.title}`);
  console.log("Sheets:");
  for (const sheetName of spreadsheet.sheets) {
    console.log(`- ${sheetName}`);
  }
}

main().catch((error) => {
  console.error("連線失敗");
  console.error(error.message);
  process.exit(1);
});
