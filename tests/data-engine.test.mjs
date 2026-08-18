import assert from "node:assert/strict";
import test from "node:test";
import * as XLSX from "xlsx";
import { applyFilters, parseWorkbook, summarize } from "../src/dataEngine.js";

function fixture() {
  const sheet = XLSX.utils.aoa_to_sheet([
    ["區域", "AC", "餐廳", "Code", "序號", "電訪日期", "同意受訪", "用餐動機", "訂餐日期", "消費金額", "成功拜訪\n訂餐日期 > 電訪日期"],
    ["北一區", "甲", "台北店", "B001", 1, new Date(2026, 6, 1), "V", "請愙", new Date(2026, 6, 2), 5000, "V"],
    ["北一區", "甲", "台北店", "B001", 2, new Date(2026, 6, 3), "V", "其他", new Date(2026, 6, 3), 3000, ""],
    ["南一區", "乙", "台南店", "B002", 1, new Date(2026, 6, 4), "", "聚餐", "", "", ""],
    ["北一區", "甲", "台北店", "B001", 2, new Date(2026, 6, 3), "V", "其他", new Date(2026, 6, 3), 3000, ""],
  ]);
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([["說明"]]), "說明");
  XLSX.utils.book_append_sheet(book, sheet, "all");
  return XLSX.write(book, { type: "array", bookType: "xlsx" });
}

test("依嚴格日期規則判斷成功並辨識重複", () => {
  const data = parseWorkbook(fixture(), "fixture.xlsx");
  assert.equal(data.selectedSheet, "all");
  assert.equal(data.rows.length, 4);
  assert.equal(data.rows[0].isSuccess, true);
  assert.equal(data.rows[1].isSuccess, false);
  assert.equal(data.rows[0].motive, "請客");
  assert.equal(data.diagnostics.duplicateCount, 1);
  assert.equal(data.diagnostics.sourceMismatch, 0);
});

test("排除重複後彙總正確", () => {
  const data = parseWorkbook(fixture());
  const rows = applyFilters(data.rows, { excludeDuplicates: true });
  const summary = summarize(rows);
  assert.equal(rows.length, 3);
  assert.equal(summary.successes, 1);
  assert.equal(summary.successAmount, 5000);
});
