import fs from "node:fs";
import path from "node:path";
import { parseWorkbook, summarize } from "../src/dataEngine.js";

const source = process.argv[2];
if (!source || !fs.existsSync(source)) {
  console.error("用法：npm run validate:workbook -- <Excel檔案路徑>");
  process.exit(1);
}
const data = parseWorkbook(fs.readFileSync(source), path.basename(source));
console.log(
  JSON.stringify(
    {
      fileName: data.fileName,
      selectedSheet: data.selectedSheet,
      dateRange: data.dateRange,
      importedRows: data.rows.length,
      complete: summarize(data.rows),
      cleaned: summarize(data.rows.filter((row) => !row.isDuplicate)),
      diagnostics: data.diagnostics,
    },
    (key, value) => (key === "issueRows" || key === "headerMapping" ? undefined : value),
    2,
  ),
);
