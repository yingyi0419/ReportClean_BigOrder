import * as XLSX from "xlsx";

const FIELDS = [
  ["region", "區域", true, ["區域", "營運區域", "區"]],
  ["ac", "AC", true, ["AC", "區經理", "督導"]],
  ["store", "餐廳", true, ["餐廳", "門市", "店名", "餐廳名稱"]],
  ["code", "Code", false, ["Code", "餐廳Code", "門市代碼", "店號"]],
  ["sequence", "序號", false, ["序號", "編號"]],
  ["callDate", "電訪日期", true, ["電訪日期", "電話拜訪日期", "聯繫日期"]],
  ["caller", "電訪人員", false, ["電訪人員", "拜訪人員"]],
  ["unreached", "未接通", false, ["未接通"]],
  ["rejected", "拒絕", false, ["拒絕", "拒訪"]],
  ["consented", "同意受訪", false, ["同意受訪", "完成受訪"]],
  ["customerType", "客戶類型", false, ["新或舊客或陌生客", "客戶類型", "新舊客"]],
  ["target", "電訪對象", false, ["電訪對象", "拜訪對象", "客戶名稱"]],
  ["attribute", "屬性", false, ["屬性", "客戶屬性"]],
  ["pkMember", "PK會員", false, ["是否為PK會員", "PK會員"]],
  ["lineMember", "LINE會員", false, ["是否為餐廳LINE會員", "LINE會員"]],
  ["motive", "用餐動機", false, ["用餐動機", "訂餐動機"]],
  ["occasion", "多人用餐時機", false, ["多人用餐時機", "用餐時機"]],
  ["budget", "多人用餐預算", false, ["多人用餐預算", "預算"]],
  ["partySize", "多人用餐人數", false, ["多人用餐人數", "用餐人數"]],
  ["preference", "餐點偏好", false, ["喜愛套餐或桶餐或皆可", "餐點偏好"]],
  ["faceDate", "面對面拜訪日期", false, ["面對面拜訪日期", "面訪日期"]],
  ["faceFocus", "面對面拜訪重點", false, ["面對面拜訪重點", "面訪重點"]],
  ["orderDate", "訂餐日期", true, ["訂餐日期", "訂單日期", "消費日期"]],
  ["amount", "消費金額", true, ["消費金額", "訂單金額", "業績", "金額"]],
  [
    "sourceSuccess",
    "來源成功拜訪",
    false,
    ["成功拜訪訂餐日期>電訪日期", "成功拜訪訂餐日期＞電訪日期", "成功拜訪"],
  ],
];

const compact = (value) =>
  String(value ?? "")
    .normalize("NFKC")
    .replace(/[\s\r\n\t_：:（）()【】[\]「」]/g, "")
    .trim()
    .toLowerCase();
const text = (value) =>
  value == null ? "" : String(value).normalize("NFKC").replace(/\s+/g, " ").trim();
const blank = (value) => value == null || text(value) === "";
const aliasMap = new Map();
FIELDS.forEach(([key, , , aliases]) =>
  aliases.forEach((alias) => aliasMap.set(compact(alias), key)),
);

export function checked(value) {
  if (value === true || value === 1) return true;
  return ["v", "✓", "✔", "是", "yes", "true", "1"].includes(text(value).toLowerCase());
}

const localDate = (year, month, day) => {
  const value = new Date(Number(year), Number(month) - 1, Number(day));
  if (
    Number.isNaN(value.getTime()) ||
    value.getFullYear() !== Number(year) ||
    value.getMonth() !== Number(month) - 1 ||
    value.getDate() !== Number(day)
  )
    return null;
  return value;
};

export function parseDate(value) {
  if (blank(value)) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime()))
    return localDate(value.getFullYear(), value.getMonth() + 1, value.getDate());
  if (typeof value === "number" && Number.isFinite(value)) {
    const utc = new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000);
    return localDate(utc.getUTCFullYear(), utc.getUTCMonth() + 1, utc.getUTCDate());
  }
  const match = text(value).match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
  return match ? localDate(match[1], match[2], match[3]) : null;
}

export function dateKey(value) {
  const valueDate = value instanceof Date ? value : parseDate(value);
  if (!valueDate) return "";
  return `${valueDate.getFullYear()}-${String(valueDate.getMonth() + 1).padStart(
    2,
    "0",
  )}-${String(valueDate.getDate()).padStart(2, "0")}`;
}

export function parseAmount(value) {
  if (blank(value)) return null;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const source = text(value);
  const negative = /^\(.*\)$/.test(source);
  const parsed = Number(source.replace(/[,$NTD新臺台幣元\s()]/gi, ""));
  return Number.isFinite(parsed) ? (negative ? -parsed : parsed) : null;
}

function inspectHeader(row) {
  const map = {};
  row.forEach((value, index) => {
    const key = aliasMap.get(compact(value));
    if (key && map[key] == null) map[key] = index;
  });
  const required = FIELDS.filter(([key, , isRequired]) => isRequired && map[key] != null)
    .length;
  const optional = FIELDS.filter(([key, , isRequired]) => !isRequired && map[key] != null)
    .length;
  return { map, required, optional, score: required * 10 + optional };
}

function selectSheet(workbook) {
  let best = null;
  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
      defval: null,
      blankrows: true,
    });
    rows.slice(0, 30).forEach((row, headerIndex) => {
      const result = inspectHeader(row);
      const candidate = {
        sheetName,
        rows,
        headerIndex,
        ...result,
        score: result.score + (compact(sheetName) === "all" ? 5 : 0),
      };
      if (!best || candidate.score > best.score) best = candidate;
    });
  });
  return best;
}

function uniqueHeaders(row) {
  const seen = new Map();
  return row.map((value, index) => {
    const base = text(value) || `未命名欄位_${index + 1}`;
    const count = (seen.get(base) ?? 0) + 1;
    seen.set(base, count);
    return count === 1 ? base : `${base}_${count}`;
  });
}

const rawValue = (row, map, key) => (map[key] == null ? null : row[map[key]]);

function normalizeRow(row, rowNumber, headers, map, sheetName) {
  const raw = {};
  headers.forEach((header, index) => {
    raw[header] = row[index] ?? null;
  });
  const callRaw = rawValue(row, map, "callDate");
  const orderRaw = rawValue(row, map, "orderDate");
  const amountRaw = rawValue(row, map, "amount");
  const callDate = parseDate(callRaw);
  const orderDate = parseDate(orderRaw);
  const amount = parseAmount(amountRaw);
  const issues = [];
  const result = {
    id: `${sheetName}-${rowNumber}`,
    rowNumber,
    region: text(rawValue(row, map, "region")),
    ac: text(rawValue(row, map, "ac")),
    store: text(rawValue(row, map, "store")),
    code: text(rawValue(row, map, "code")),
    sequence: text(rawValue(row, map, "sequence")),
    callDate,
    callDateKey: dateKey(callDate),
    caller: text(rawValue(row, map, "caller")),
    unreached: checked(rawValue(row, map, "unreached")),
    rejected: checked(rawValue(row, map, "rejected")),
    consented: checked(rawValue(row, map, "consented")),
    customerType: text(rawValue(row, map, "customerType")) || "未分類",
    target: text(rawValue(row, map, "target")),
    attribute: text(rawValue(row, map, "attribute")),
    pkMember: text(rawValue(row, map, "pkMember")),
    lineMember: text(rawValue(row, map, "lineMember")),
    motive: text(rawValue(row, map, "motive"))
      .replace(/^請愙$/, "請客")
      .replace(/^其它$/, "其他"),
    occasion: text(rawValue(row, map, "occasion")),
    budget: parseAmount(rawValue(row, map, "budget")),
    partySize: text(rawValue(row, map, "partySize")),
    preference: text(rawValue(row, map, "preference")),
    faceDate: parseDate(rawValue(row, map, "faceDate")),
    faceFocus: text(rawValue(row, map, "faceFocus")),
    orderDate,
    orderDateKey: dateKey(orderDate),
    amount,
    isSuccess: Boolean(callDate && orderDate && orderDate > callDate),
    sourceSuccess: checked(rawValue(row, map, "sourceSuccess")),
    issues,
    isDuplicate: false,
    raw,
  };
  if (!result.region) issues.push("缺少區域");
  if (!result.ac) issues.push("缺少 AC");
  if (!result.store) issues.push("缺少餐廳");
  if (!callDate) issues.push(blank(callRaw) ? "缺少電訪日期" : "電訪日期無法辨識");
  if (!blank(orderRaw) && !orderDate) issues.push("訂餐日期無法辨識");
  if (!blank(amountRaw) && amount == null) issues.push("消費金額無法辨識");
  if (result.isSuccess && amount == null) issues.push("成功拜訪缺少消費金額");
  if ([result.unreached, result.rejected, result.consented].filter(Boolean).length > 1)
    issues.push("接通結果重複勾選");
  result.duplicateKey =
    result.callDateKey && result.sequence && (result.code || result.store)
      ? `${result.code || result.store}|${result.sequence}|${result.callDateKey}`
      : "";
  return result;
}

export function parseWorkbook(buffer, fileName = "已上傳活頁簿.xlsx") {
  let workbook;
  try {
    workbook = XLSX.read(buffer, {
      type: "array",
      cellDates: false,
      cellNF: true,
      dense: true,
    });
  } catch {
    throw new Error("無法讀取活頁簿，請確認檔案未損毀且格式為 .xlsx 或 .xls。");
  }
  const selected = selectSheet(workbook);
  if (!selected) throw new Error("找不到可分析的工作表。");
  const missing = FIELDS.filter(
    ([key, , required]) => required && selected.map[key] == null,
  );
  if (missing.length)
    throw new Error(`找不到必要資料行：${missing.map(([, label]) => label).join("、")}`);

  const sourceHeader = selected.rows[selected.headerIndex];
  const headers = uniqueHeaders(sourceHeader);
  const rows = [];
  let ignored = 0;
  selected.rows.slice(selected.headerIndex + 1).forEach((row, offset) => {
    if (!row.some((value) => !blank(value))) return;
    if (
      !["region", "ac", "store", "code", "callDate", "orderDate"].some(
        (key) => !blank(rawValue(row, selected.map, key)),
      )
    ) {
      ignored += 1;
      return;
    }
    rows.push(
      normalizeRow(
        row,
        selected.headerIndex + offset + 2,
        headers,
        selected.map,
        selected.sheetName,
      ),
    );
  });
  if (!rows.length) throw new Error("已找到欄名，但沒有可分析的明細資料列。");

  const seen = new Set();
  rows.forEach((row) => {
    if (!row.duplicateKey) return;
    if (seen.has(row.duplicateKey)) {
      row.isDuplicate = true;
      row.issues.push("疑似重複資料");
    } else seen.add(row.duplicateKey);
  });

  const callDates = rows.map((row) => row.callDateKey).filter(Boolean).sort();
  const sourceSuccessAvailable = selected.map.sourceSuccess != null;
  const headerMapping = FIELDS.map(([key, label, required]) => ({
    key,
    label,
    required,
    actual: selected.map[key] == null ? "" : headers[selected.map[key]],
    found: selected.map[key] != null,
  }));
  const diagnostics = {
    sheetNames: workbook.SheetNames,
    selectedSheet: selected.sheetName,
    headerRow: selected.headerIndex + 1,
    importedRows: rows.length,
    sourceColumns: headers.length,
    ignored,
    duplicateCount: rows.filter((row) => row.isDuplicate).length,
    invalidCallDate: rows.filter((row) => !row.callDate).length,
    missingOrderDate: rows.filter((row) => !row.orderDate).length,
    invalidAmount: rows.filter((row) => row.issues.includes("消費金額無法辨識"))
      .length,
    successMissingAmount: rows.filter((row) =>
      row.issues.includes("成功拜訪缺少消費金額"),
    ).length,
    multipleStatus: rows.filter((row) =>
      row.issues.includes("接通結果重複勾選"),
    ).length,
    sourceSuccessAvailable,
    sourceMismatch: sourceSuccessAvailable
      ? rows.filter((row) => row.sourceSuccess !== row.isSuccess).length
      : 0,
    issueRows: rows.filter((row) => row.issues.length),
    headerMapping,
  };
  return {
    fileName,
    selectedSheet: selected.sheetName,
    rows,
    headers,
    diagnostics,
    dateRange: { min: callDates[0] || "", max: callDates.at(-1) || "" },
  };
}

export async function parseFile(file) {
  return parseWorkbook(await file.arrayBuffer(), file.name);
}

export function applyFilters(rows, filters = {}) {
  const search = text(filters.search).toLowerCase();
  return rows.filter((row) => {
    if (filters.excludeDuplicates && row.isDuplicate) return false;
    if (filters.start && row.callDateKey < filters.start) return false;
    if (filters.end && row.callDateKey > filters.end) return false;
    if (filters.region && row.region !== filters.region) return false;
    if (filters.ac && row.ac !== filters.ac) return false;
    if (filters.store && row.store !== filters.store) return false;
    if (filters.success === "yes" && !row.isSuccess) return false;
    if (filters.success === "no" && row.isSuccess) return false;
    if (
      search &&
      ![
        row.region,
        row.ac,
        row.store,
        row.code,
        row.caller,
        row.customerType,
        row.target,
        row.motive,
      ]
        .join(" ")
        .toLowerCase()
        .includes(search)
    )
      return false;
    return true;
  });
}

export function summarize(rows) {
  const successful = rows.filter((row) => row.isSuccess);
  const withAmount = successful.filter((row) => row.amount != null);
  const successAmount = withAmount.reduce((sum, row) => sum + row.amount, 0);
  const consented = rows.filter((row) => row.consented).length;
  return {
    visits: rows.length,
    successes: successful.length,
    successRate: rows.length ? successful.length / rows.length : 0,
    successAmount,
    averageAmount: withAmount.length ? successAmount / withAmount.length : 0,
    amountRows: withAmount.length,
    missingAmount: successful.length - withAmount.length,
    consented,
    consentRate: rows.length ? consented / rows.length : 0,
    unreached: rows.filter((row) => row.unreached).length,
    rejected: rows.filter((row) => row.rejected).length,
  };
}

export function groupBy(rows, key) {
  const groups = new Map();
  rows.forEach((row) => {
    const name = text(row[key]) || "未分類";
    if (!groups.has(name)) groups.set(name, []);
    groups.get(name).push(row);
  });
  return [...groups].map(([name, values]) => ({ name, ...summarize(values) }));
}

function weekKey(date) {
  const start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  start.setDate(start.getDate() - (start.getDay() || 7) + 1);
  return dateKey(start);
}

export function weekly(rows) {
  const groups = new Map();
  rows.forEach((row) => {
    if (!row.callDate) return;
    const key = weekKey(row.callDate);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  });
  return [...groups]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([key, values]) => ({
      week: key,
      name: `${Number(key.slice(5, 7))}/${Number(key.slice(8))}`,
      ...summarize(values),
    }));
}

function sheetFromRows(headers, body) {
  const sheet = XLSX.utils.aoa_to_sheet([headers, ...body]);
  sheet["!autofilter"] = { ref: sheet["!ref"] };
  sheet["!cols"] = headers.map((header) => ({
    wch: Math.min(Math.max(String(header).length * 2, 11), 30),
  }));
  return sheet;
}

const amountBand = (amount) => {
  if (amount == null) return "金額缺漏／異常";
  if (amount < 5000) return "未滿 5,000";
  if (amount < 10000) return "5,000–9,999";
  if (amount < 20000) return "10,000–19,999";
  if (amount < 50000) return "20,000–49,999";
  if (amount < 100000) return "50,000–99,999";
  return "100,000 以上";
};

function analysisRows(rows, key, normalize = (value) => text(value) || "未分類") {
  const groups = new Map();
  rows.forEach((row) => {
    const name = normalize(row[key], row);
    if (!groups.has(name)) groups.set(name, []);
    groups.get(name).push(row);
  });
  return [...groups]
    .map(([name, values]) => ({ name, ...summarize(values) }))
    .sort((a, b) => b.successAmount - a.successAmount || b.successes - a.successes);
}

function appendAnalysisSheet(book, sheetName, data) {
  const sheet = sheetFromRows(
    ["分類", "拜訪筆數", "成功筆數", "成功率", "成功消費金額", "成功平均金額", "有效金額筆數", "成功缺漏金額筆數"],
    data.map((item) => [item.name, item.visits, item.successes, item.successRate, item.successAmount, item.averageAmount, item.amountRows, item.missingAmount]),
  );
  for (let row = 2; row <= data.length + 1; row += 1) {
    if (sheet[`D${row}`]) sheet[`D${row}`].z = "0.0%";
    ["E", "F"].forEach((column) => {
      if (sheet[`${column}${row}`]) sheet[`${column}${row}`].z = "#,##0";
    });
  }
  XLSX.utils.book_append_sheet(book, sheet, sheetName);
}

export function exportWorkbook({ dataset, rows, filters, complete = false }) {
  const target = complete ? dataset.rows : rows;
  const summary = summarize(target);
  const book = XLSX.utils.book_new();
  const summarySheet = XLSX.utils.aoa_to_sheet([
    ["KFC 大單拜訪分析｜匯出摘要", ""],
    ["來源檔案", dataset.fileName],
    ["採用工作表", dataset.selectedSheet],
    ["匯出範圍", complete ? "完整清理資料" : "目前篩選結果"],
    ["日期起", filters.start || "不限"],
    ["日期迄", filters.end || "不限"],
    ["區域", filters.region || "全部"],
    ["AC", filters.ac || "全部"],
    ["餐廳", filters.store || "全部"],
    ["排除疑似重複", filters.excludeDuplicates ? "是" : "否"],
    ["", ""],
    ["有效拜訪", summary.visits],
    ["成功拜訪", summary.successes],
    ["成功率", summary.successRate],
    ["成功消費金額", summary.successAmount],
    ["成功平均金額", summary.averageAmount],
    ["同意受訪率", summary.consentRate],
    ["成功但缺金額", summary.missingAmount],
    ["成功定義", "訂餐日期 > 電訪日期（嚴格大於）"],
    ["資料處理", "瀏覽器本機解析，未上傳伺服器"],
  ]);
  summarySheet["!cols"] = [{ wch: 24 }, { wch: 48 }];
  summarySheet.B14.z = "0.0%";
  summarySheet.B17.z = "0.0%";
  XLSX.utils.book_append_sheet(book, summarySheet, "分析摘要");

  const extras = [
    "清理_區域",
    "清理_AC",
    "清理_餐廳",
    "清理_電訪日期",
    "清理_訂餐日期",
    "清理_消費金額",
    "系統判斷_成功拜訪",
    "疑似重複",
    "資料品質註記",
  ];
  const body = target.map((row) => [
    ...dataset.headers.map((header) => row.raw[header] ?? null),
    row.region,
    row.ac,
    row.store,
    row.callDate || null,
    row.orderDate || null,
    row.amount,
    row.isSuccess ? "V" : "",
    row.isDuplicate ? "V" : "",
    row.issues.join("；"),
  ]);
  XLSX.utils.book_append_sheet(
    book,
    sheetFromRows([...dataset.headers, ...extras], body),
    complete ? "完整清理資料" : "篩選明細",
  );
  if (complete) {
    appendAnalysisSheet(book, "新舊客分析", analysisRows(target, "customerType"));
    appendAnalysisSheet(book, "屬性分析", analysisRows(target, "attribute"));
    appendAnalysisSheet(book, "PK會員分析", analysisRows(target, "pkMember"));
    appendAnalysisSheet(book, "LINE會員分析", analysisRows(target, "lineMember"));
    appendAnalysisSheet(book, "用餐動機分析", analysisRows(target, "motive"));
    appendAnalysisSheet(book, "金額區間分析", analysisRows(target, "amount", (value) => amountBand(value)));

    const topOrders = target
      .filter((row) => row.isSuccess && row.amount != null)
      .sort((a, b) => b.amount - a.amount);
    const topSheet = sheetFromRows(
      ["排名", "成功訂單金額", "區域", "AC", "餐廳", "客戶／電訪對象", "電訪日期", "訂餐日期", "成功天數差", "來源資料列"],
      topOrders.map((row, index) => [
        index + 1, row.amount, row.region, row.ac, row.store, row.target,
        row.callDate || null, row.orderDate || null,
        row.callDate && row.orderDate ? Math.round((row.orderDate - row.callDate) / 86400000) : null,
        row.rowNumber,
      ]),
    );
    for (let row = 2; row <= topOrders.length + 1; row += 1) {
      if (topSheet[`B${row}`]) topSheet[`B${row}`].z = "#,##0";
      ["G", "H"].forEach((column) => {
        if (topSheet[`${column}${row}`]) topSheet[`${column}${row}`].z = "yyyy/mm/dd";
      });
    }
    XLSX.utils.book_append_sheet(book, topSheet, "最佳成功訂單明細");
  }
  XLSX.utils.book_append_sheet(
    book,
    sheetFromRows(
      ["診斷項目", "結果"],
      [
        ["匯入明細", dataset.diagnostics.importedRows],
        ["疑似重複", dataset.diagnostics.duplicateCount],
        ["無效電訪日期", dataset.diagnostics.invalidCallDate],
        ["缺少訂餐日期", dataset.diagnostics.missingOrderDate],
        ["成功但缺少金額", dataset.diagnostics.successMissingAmount],
        ["來源成功欄差異", dataset.diagnostics.sourceMismatch],
        ...dataset.diagnostics.headerMapping.map((item) => [
          `${item.label}${item.required ? "（必要）" : ""}`,
          item.actual || "未找到",
        ]),
      ],
    ),
    "資料診斷",
  );
  const now = new Date();
  const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(
    2,
    "0",
  )}${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(
    2,
    "0",
  )}${String(now.getMinutes()).padStart(2, "0")}`;
  XLSX.writeFile(
    book,
    `KFC_大單拜訪分析_${complete ? "完整清理資料" : "篩選結果"}_${stamp}.xlsx`,
    { compression: true },
  );
}
