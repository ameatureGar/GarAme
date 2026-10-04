const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");

function colRow(ref) {
  const m = ref.match(/^([A-Z]+)(\d+)$/);
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col, row: Number(m[2]) };
}

function parseSharedStrings(xml) {
  const strings = [];
  const siRe = /<si>([\s\S]*?)<\/si>/g;
  let m;
  while ((m = siRe.exec(xml))) {
    const texts = [...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((x) => decode(x[1]));
    strings.push(texts.join(""));
  }
  return strings;
}

function decode(s) {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)));
}

function parseSheet(xml, strings) {
  const rows = new Map();
  const cellRe = /<c r="([A-Z]+\d+)"([^>]*)(?:\/>|>([\s\S]*?)<\/c>)/g;
  let m;
  while ((m = cellRe.exec(xml))) {
    const { col, row } = colRow(m[1]);
    const attrs = m[2] || "";
    const inner = m[3] || "";
    const t = (attrs.match(/ t="([^"]+)"/) || [])[1];
    const v = (inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
    const is = (inner.match(/<is>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>/) || [])[1];
    let val = "";
    if (t === "s") val = strings[Number(v)] ?? "";
    else if (t === "inlineStr") val = is ? decode(is) : "";
    else if (t === "b") val = v === "1";
    else if (v != null && v !== "") val = Number(v);
    else val = "";
    if (!rows.has(row)) rows.set(row, {});
    rows.get(row)[col] = val;
  }
  return rows;
}

function extractXlsx(xlsxPath, destDir) {
  const zipPath = path.join(destDir, "book.zip");
  fs.copyFileSync(xlsxPath, zipPath);
  execFileSync("tar", ["-xf", zipPath, "-C", destDir], { windowsHide: true });
}

function loadWorkbook(xlsxPath) {
  const destDir = fs.mkdtempSync(path.join(os.tmpdir(), "kshma-xlsx-"));
  try {
    extractXlsx(xlsxPath, destDir);
    const sstPath = path.join(destDir, "xl", "sharedStrings.xml");
    const strings = fs.existsSync(sstPath)
      ? parseSharedStrings(fs.readFileSync(sstPath, "utf8"))
      : [];
    const sheetPath = path.join(destDir, "xl", "worksheets", "sheet1.xml");
    if (!fs.existsSync(sheetPath)) throw new Error("The workbook does not contain Sheet1.");
    const rows = parseSheet(fs.readFileSync(sheetPath, "utf8"), strings);
    return {
      cell(r, c) {
        const row = rows.get(r);
        if (!row) return "";
        const v = row[c];
        return v == null ? "" : v;
      },
      rows,
      maxRow: rows.size ? Math.max(...rows.keys()) : 0,
    };
  } finally {
    fs.rmSync(destDir, { recursive: true, force: true });
  }
}

module.exports = { loadWorkbook };
