import {
  buildExportBaseName,
  ensureExtension,
  sanitizeFilename
} from "../filename.js";
import { createZipBlob } from "../zip.js";

const XLSX_MAX_COLUMNS = 16384;
const XLSX_MAX_ROWS = 1048576;
const XLSX_SHEET_NAME_MAX = 31;

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const XLSX_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>`;

const XLSX_ROOT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship
    Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"
  />
</Relationships>`;

function sanitizeXmlChars(value) {
  return String(value || "").replace(/[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD]/g, "");
}

function escapeXml(value) {
  return sanitizeXmlChars(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function stringifyUnknown(value) {
  const seen = new WeakSet();
  try {
    return JSON.stringify(value, (_key, current) => {
      if (typeof current === "bigint") return current.toString();
      if (!current || typeof current !== "object") return current;
      if (seen.has(current)) return "[Circular]";
      seen.add(current);
      return current;
    });
  } catch {
    return String(value);
  }
}

function normalizeCellValue(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean" || typeof value === "bigint") {
    return String(value);
  }
  if (value instanceof Date) return value.toISOString();
  if (Array.isArray(value)) {
    const allPrimitive = value.every((item) =>
      item === null ||
      item === undefined ||
      typeof item === "string" ||
      typeof item === "number" ||
      typeof item === "boolean"
    );
    if (allPrimitive) {
      return value.map((item) => (item === null || item === undefined ? "" : String(item))).join("; ");
    }
    return stringifyUnknown(value);
  }
  return stringifyUnknown(value);
}

function normalizeString(value, fallback) {
  const normalized = String(value || "").trim();
  return normalized || fallback;
}

function resolveBitableBaseTitle(payload, fallback = "bitable-export") {
  const rawTitle = normalizeString(payload?.title, "");
  const firstTable = Array.isArray(payload?.bitableData?.tables)
    ? payload.bitableData.tables[0]
    : null;
  const firstTableName = normalizeString(firstTable?.tableName || firstTable?.name, "");
  const looksGenericTitle = /^(导出任务|bitable(\s+export)?|表格)$/i.test(rawTitle);
  const preferredTitle =
    (firstTableName && (looksGenericTitle || !rawTitle) ? firstTableName : rawTitle) ||
    firstTableName ||
    fallback;
  return preferredTitle;
}

function extractRecordFields(record) {
  if (!record || typeof record !== "object") {
    return {};
  }
  if (record.fields && typeof record.fields === "object") {
    return record.fields;
  }
  if (record.cellValueMap && typeof record.cellValueMap === "object") {
    return record.cellValueMap;
  }
  const fallback = {};
  for (const [key, value] of Object.entries(record)) {
    if (key === "rowIndex" || key === "recordId" || key === "id") continue;
    fallback[key] = value;
  }
  return fallback;
}

function normalizeTableRows(table) {
  const records = Array.isArray(table?.records) ? table.records : [];
  return records.map((record, index) => {
    const fields = extractRecordFields(record);
    const rowIndexValue = Number(record?.rowIndex);
    const rowIndex = Number.isFinite(rowIndexValue) && rowIndexValue > 0 ? rowIndexValue : index + 1;
    const recordId = normalizeString(record?.recordId || record?.id || "", "");
    return {
      rowIndex,
      recordId,
      fields
    };
  });
}

function collectOrderedFieldColumns(table, rows) {
  const ordered = [];
  const seen = new Set();
  const push = (name) => {
    const key = normalizeString(name, "");
    if (!key || seen.has(key)) return;
    seen.add(key);
    ordered.push(key);
  };

  const tableFields = Array.isArray(table?.fields) ? table.fields : [];
  for (const field of tableFields) {
    push(field?.name || field?.id);
  }
  for (const row of rows) {
    for (const key of Object.keys(row.fields || {})) {
      push(key);
    }
  }
  return ordered.slice(0, Math.max(1, XLSX_MAX_COLUMNS - 2));
}

function normalizeTableData(table, index) {
  const rows = normalizeTableRows(table);
  const fieldColumns = collectOrderedFieldColumns(table, rows);
  const tableName = normalizeString(table?.tableName || table?.name || "", `表格${index + 1}`);
  const columns = ["序号", ...fieldColumns];
  const rowValues = rows.map((row) => {
    const values = [row.rowIndex];
    for (const fieldName of fieldColumns) {
      values.push(normalizeCellValue(row.fields?.[fieldName]));
    }
    return values;
  });
  return {
    tableName,
    columns,
    rows: rowValues
  };
}

function normalizeAllTables(bitableData) {
  const tables = Array.isArray(bitableData?.tables) ? bitableData.tables : [];
  return tables.map((table, index) => normalizeTableData(table, index));
}

function escapeCsvCell(value) {
  const text = normalizeCellValue(value);
  const shouldQuote = /[",\r\n]/.test(text) || /^\s|\s$/.test(text);
  if (!shouldQuote) return text;
  return `"${text.replace(/"/g, "\"\"")}"`;
}

function buildCsvText(tableData) {
  const lines = [];
  lines.push(tableData.columns.map((header) => escapeCsvCell(header)).join(","));
  for (const row of tableData.rows) {
    lines.push(row.map((cell) => escapeCsvCell(cell)).join(","));
  }
  return `\uFEFF${lines.join("\r\n")}\r\n`;
}

function sanitizeSheetName(input, index) {
  const fallback = `Sheet${index + 1}`;
  let name = normalizeString(input, fallback)
    .replace(/[\\/*?:[\]]/g, "_")
    .replace(/^'+|'+$/g, "")
    .trim();
  if (!name) {
    name = fallback;
  }
  const chars = Array.from(name);
  if (chars.length > XLSX_SHEET_NAME_MAX) {
    name = chars.slice(0, XLSX_SHEET_NAME_MAX).join("");
  }
  return name || fallback;
}

function makeUniqueSheetNames(tableDataList) {
  const used = new Set();
  return tableDataList.map((tableData, index) => {
    const base = sanitizeSheetName(tableData.tableName, index);
    let name = base;
    let counter = 2;
    while (used.has(name.toLowerCase())) {
      const suffix = `_${counter}`;
      const maxHead = Math.max(1, XLSX_SHEET_NAME_MAX - suffix.length);
      const head = Array.from(base).slice(0, maxHead).join("");
      name = `${head}${suffix}`;
      counter += 1;
    }
    used.add(name.toLowerCase());
    return name;
  });
}

function toColumnLetters(colIndex1Based) {
  let value = colIndex1Based;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function buildWorksheetXml(tableData) {
  const columns = tableData.columns.slice(0, XLSX_MAX_COLUMNS);
  const maxDataRows = Math.max(0, XLSX_MAX_ROWS - 1);
  const dataRows = tableData.rows.slice(0, maxDataRows);
  const rowXml = [];

  const headerCells = columns
    .map(
      (header, index) =>
        `<c r="${toColumnLetters(index + 1)}1" t="inlineStr"><is><t xml:space="preserve">${escapeXml(
          header
        )}</t></is></c>`
    )
    .join("");
  rowXml.push(`<row r="1">${headerCells}</row>`);

  for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex += 1) {
    const excelRow = rowIndex + 2;
    const row = dataRows[rowIndex];
    let cells = "";
    for (let colIndex = 0; colIndex < columns.length; colIndex += 1) {
      const cellRef = `${toColumnLetters(colIndex + 1)}${excelRow}`;
      const text = normalizeCellValue(row[colIndex]);
      cells += `<c r="${cellRef}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(
        text
      )}</t></is></c>`;
    }
    rowXml.push(`<row r="${excelRow}">${cells}</row>`);
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    ${rowXml.join("")}
  </sheetData>
</worksheet>`;
}

function buildWorkbookXml(sheetNames) {
  const sheetsXml = sheetNames
    .map(
      (sheetName, index) =>
        `<sheet name="${escapeXml(sheetName)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`
    )
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheetsXml}</sheets>
</workbook>`;
}

function buildWorkbookRelsXml(sheetCount) {
  const items = [];
  for (let i = 1; i <= sheetCount; i += 1) {
    items.push(
      `<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i}.xml"/>`
    );
  }
  items.push(
    `<Relationship Id="rId${sheetCount + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
  );
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${items.join("")}
</Relationships>`;
}

function buildContentTypesXml(sheetCount) {
  const worksheetOverrides = [];
  for (let i = 1; i <= sheetCount; i += 1) {
    worksheetOverrides.push(
      `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    );
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  ${worksheetOverrides.join("")}
</Types>`;
}

async function buildXlsxBlob(tableDataList) {
  const sheetNames = makeUniqueSheetNames(tableDataList);
  const files = [
    { name: "[Content_Types].xml", data: buildContentTypesXml(tableDataList.length) },
    { name: "_rels/.rels", data: XLSX_ROOT_RELS_XML },
    { name: "xl/workbook.xml", data: buildWorkbookXml(sheetNames) },
    {
      name: "xl/_rels/workbook.xml.rels",
      data: buildWorkbookRelsXml(tableDataList.length)
    },
    { name: "xl/styles.xml", data: XLSX_STYLES_XML }
  ];

  for (let i = 0; i < tableDataList.length; i += 1) {
    files.push({
      name: `xl/worksheets/sheet${i + 1}.xml`,
      data: buildWorksheetXml(tableDataList[i])
    });
  }
  const zipBlob = await createZipBlob(files);
  return new Blob([zipBlob], { type: XLSX_MIME });
}

function buildUniqueCsvName(baseName, index, used) {
  const safeBase = sanitizeFilename(baseName, `table-${index + 1}`, 80);
  let fileName = ensureExtension(safeBase, "csv");
  let seed = 2;
  while (used.has(fileName.toLowerCase())) {
    fileName = ensureExtension(`${safeBase}-${seed}`, "csv");
    seed += 1;
  }
  used.add(fileName.toLowerCase());
  return fileName;
}

export async function buildBitableJsonExport(payload) {
  const baseName = buildExportBaseName(
    resolveBitableBaseTitle(payload, "bitable-export"),
    "bitable-export"
  );
  return {
    blob: new Blob([JSON.stringify(payload?.bitableData || {}, null, 2)], {
      type: "application/json;charset=utf-8"
    }),
    filename: ensureExtension(baseName, "json"),
    exportType: "BITABLE_JSON"
  };
}

export async function buildBitableCsvExport(payload) {
  const tableDataList = normalizeAllTables(payload?.bitableData);
  if (!tableDataList.length) {
    throw new Error("未读取到可导出的 Bitable 表格数据");
  }

  const baseName = buildExportBaseName(
    resolveBitableBaseTitle(payload, "bitable-export"),
    "bitable-export"
  );
  if (tableDataList.length === 1) {
    const csvText = buildCsvText(tableDataList[0]);
    return {
      blob: new Blob([csvText], { type: "text/csv;charset=utf-8" }),
      filename: ensureExtension(baseName, "csv"),
      exportType: "BITABLE_CSV",
      tableCount: 1
    };
  }

  const usedFileNames = new Set();
  const csvFiles = tableDataList.map((tableData, index) => ({
    name: buildUniqueCsvName(tableData.tableName, index, usedFileNames),
    data: buildCsvText(tableData)
  }));
  return {
    blob: await createZipBlob(csvFiles),
    filename: ensureExtension(baseName, "zip"),
    exportType: "BITABLE_CSV",
    tableCount: tableDataList.length
  };
}

export async function buildBitableXlsxExport(payload) {
  const tableDataList = normalizeAllTables(payload?.bitableData);
  if (!tableDataList.length) {
    throw new Error("未读取到可导出的 Bitable 表格数据");
  }
  const baseName = buildExportBaseName(
    resolveBitableBaseTitle(payload, "bitable-export"),
    "bitable-export"
  );
  const xlsxBlob = await buildXlsxBlob(tableDataList);
  return {
    blob: xlsxBlob,
    filename: ensureExtension(baseName, "xlsx"),
    exportType: "BITABLE_XLSX",
    tableCount: tableDataList.length
  };
}
