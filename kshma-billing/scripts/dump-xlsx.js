const fs = require("fs");
const path = require("path");
const { inflateSync } = require("zlib");

const sheetPath = path.join(__dirname, "..", "tmp-xlsx", "unz", "xl", "worksheets", "sheet1.xml");
const sstPath = path.join(__dirname, "..", "tmp-xlsx", "unz", "xl", "sharedStrings.xml");

const sstXml = fs.readFileSync(sstPath, "utf8");
const strings = [];
const siRe = /<si>([\s\S]*?)<\/si>/g;
let m;
while ((m = siRe.exec(sstXml))) {
  const texts = [...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((x) =>
    x[1]
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
  );
  strings.push(texts.join(""));
}

function colRow(ref) {
  const mm = ref.match(/^([A-Z]+)(\d+)$/);
  let col = 0;
  for (const ch of mm[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col, row: Number(mm[2]), ref };
}

const xml = fs.readFileSync(sheetPath, "utf8");
const rows = new Map();
const cellRe = /<c r="([A-Z]+\d+)"([^>]*)(?:\/>|>([\s\S]*?)<\/c>)/g;
while ((m = cellRe.exec(xml))) {
  const { col, row, ref } = colRow(m[1]);
  const attrs = m[2] || "";
  const inner = m[3] || "";
  const t = (attrs.match(/ t="([^"]+)"/) || [])[1];
  let val = "";
  const v = (inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
  const is = (inner.match(/<is>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>/) || [])[1];
  if (t === "s") val = strings[Number(v)] ?? "";
  else if (t === "inlineStr") val = is || "";
  else if (t === "b") val = v === "1" ? "TRUE" : "FALSE";
  else val = v ?? "";
  const f = (inner.match(/<f[^>]*>([\s\S]*?)<\/f>/) || [])[1];
  if (!rows.has(row)) rows.set(row, {});
  rows.get(row)[col] = { ref, val, f, t };
}

const merged = [];
const mergeRe = /<mergeCell ref="([^"]+)"/g;
while ((m = mergeRe.exec(xml))) merged.push(m[1]);

const rowNums = [...rows.keys()].sort((a, b) => a - b);
const maxCol = Math.max(...rowNums.flatMap((r) => Object.keys(rows.get(r)).map(Number)));
const lines = [];
lines.push("MERGES: " + merged.join(", "));
lines.push("MAXCOL: " + maxCol);
for (const r of rowNums) {
  const cells = rows.get(r);
  const parts = [];
  for (let c = 1; c <= maxCol; c++) {
    const cell = cells[c];
    if (!cell) continue;
    const show = cell.f ? `${cell.val}[=${cell.f}]` : cell.val;
    if (show !== "" && show != null) parts.push(`${cell.ref}:${JSON.stringify(show)}`);
  }
  if (parts.length) lines.push(`R${r}\t` + parts.join(" | "));
}
fs.writeFileSync(path.join(__dirname, "..", "tmp-xlsx", "sheet1-dump.txt"), lines.join("\n"));
console.log("rows", rowNums.length, "written", lines.length);
