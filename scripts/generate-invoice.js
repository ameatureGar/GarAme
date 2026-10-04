const fs = require("fs");
const path = require("path");
const { loadWorkbook } = require("./xlsx");

const ROOT = path.join(__dirname, "..");
const meta = JSON.parse(fs.readFileSync(path.join(ROOT, "data", "invoice-meta.json"), "utf8"));

class WorkbookFormatError extends Error {
  constructor(message, details = []) {
    super(message);
    this.name = "WorkbookFormatError";
    this.details = details;
  }
}

function r2(n) {
  return Math.round((Number(n) + Number.EPSILON) * 100) / 100;
}
function num(v) {
  if (v === "" || v == null || v === false) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}
function money(n) {
  if (n == null || n === "") return "";
  return r2(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function qty(n, d = 2) {
  if (n == null || n === "") return "";
  return Number(n).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
}
function plain(n) {
  if (n == null || n === "") return "";
  if (typeof n === "number") {
    if (Number.isInteger(n)) return String(n);
    return String(Number(n.toFixed(4)).toString());
  }
  return String(n);
}
function esc(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
function unitForInvoice(u) {
  const t = String(u || "").trim().toLowerCase();
  if (["sq.mtr", "sqmt", "sqmts", "m2", "sq.m"].includes(t)) return "Smt";
  if (["sqft", "sft", "sq.ft"].includes(t)) return "Sft";
  if (["m3", "cum"].includes(t)) return "M3";
  if (["mt", "tons"].includes(t)) return "MT";
  if (["nos", "no"].includes(t)) return "Nos";
  return String(u || "").trim();
}

function letterheadDataUri() {
  const jpg = path.join(ROOT, "assets", "letterhead.jpg");
  const png = path.join(ROOT, "assets", "letterhead.png");
  if (fs.existsSync(jpg)) {
    return "data:image/jpeg;base64," + fs.readFileSync(jpg).toString("base64");
  }
  return "data:image/png;base64," + fs.readFileSync(png).toString("base64");
}

function rowVals(wb, r, maxCol) {
  const out = [];
  for (let c = 1; c <= maxCol; c++) out.push(wb.cell(r, c));
  return out;
}
function rowEmpty(vals) {
  return vals.every((v) => v === "" || v == null);
}

function buildModel(wb) {
  const abstractHeader = findRow(wb, 1, Math.min(wb.maxRow, 60), (r) => {
    const b = normalized(wb.cell(r, 2));
    const c = normalized(wb.cell(r, 3));
    const d = normalized(wb.cell(r, 4));
    return b.includes("description") && c.includes("unit") && d.includes("quantity");
  });
  if (!abstractHeader) {
    throw new WorkbookFormatError("The bill table could not be found.", [
      'Missing headings: "Description", "Unit", and "Quantity".',
    ]);
  }

  const items = [];
  for (let r = abstractHeader + 1; r <= Math.min(abstractHeader + 50, wb.maxRow); r++) {
    const desc = String(wb.cell(r, 2) || "").trim();
    if (!desc) {
      if (items.length) break;
      continue;
    }
    const unit = unitForInvoice(wb.cell(r, 3));
    const totalQty = num(wb.cell(r, 4));
    const totalRate = num(wb.cell(r, 5));
    const presentQtyRaw = num(wb.cell(r, 7));
    const presentRate = num(wb.cell(r, 8)) ?? totalRate;
    const presentQty = presentQtyRaw == null ? null : r2(presentQtyRaw);
    const presentAmt = presentQty == null ? 0 : r2(presentQty * presentRate);
    if (!unit || totalQty == null || totalRate == null || presentQtyRaw == null) continue;
    const sl = items.length + 1;
    items.push({
      sl,
      desc,
      unit,
      qty: presentQty,
      rate: presentRate,
      amount: presentAmt,
    });
  }
  if (!items.length) {
    throw new WorkbookFormatError("The bill table has no usable line items.", [
      "Each bill row needs a description, unit, total quantity, rate, and present quantity.",
    ]);
  }
  const taxable = r2(items.reduce((s, it) => s + it.amount, 0));
  const cgst = r2(taxable * meta.gstRate);
  const sgst = r2(taxable * meta.gstRate);
  const grand = r2(taxable + cgst + sgst);

  const headerRows = [];
  for (let r = abstractHeader + 1; r <= wb.maxRow; r++) {
    if (normalized(wb.cell(r, 2)).includes("particular") && normalized(wb.cell(r, 3)) === "nos") {
      headerRows.push(r);
    }
  }
  const blockStart = headerRows[0];
  const plasterStart = headerRows[1];
  const blockEnd = blockStart && plasterStart
    ? findLastRow(wb, blockStart, plasterStart - 1, (r) => rowContains(wb, r, ["cum"]))
    : null;
  const steelData = findRow(wb, (plasterStart || abstractHeader) + 1, wb.maxRow, (r) =>
    normalized(wb.cell(r, 2)).includes("plinth beam shuttering")
  );
  const steelStart = steelData
    ? findImmediateHeader(wb, steelData) || steelData
    : null;
  const earthData = findRow(wb, (steelStart || abstractHeader) + 1, wb.maxRow, (r) =>
    normalized(wb.cell(r, 9)) === "sft" && num(wb.cell(r, 8)) != null
  );
  const earthStart = earthData ? findImmediateColumnHeader(wb, earthData) || earthData : null;
  const plasterEnd = plasterStart && steelData
    ? findLastRow(wb, plasterStart, steelData - 1, (r) => rowContains(wb, r, ["m2"]))
    : null;
  const steelEnd = steelStart && earthStart
    ? findLastRow(wb, steelStart, earthStart - 1, (r) => rowContains(wb, r, ["tons"]))
    : null;

  const missingSections = [];
  if (!blockStart) missingSections.push('First measurement heading "PARTICULAR / NOS" is missing.');
  if (!plasterStart) missingSections.push('Second measurement heading "PARTICULAR / NOS" is missing.');
  if (!blockEnd) missingSections.push('Block work total in "Cum" is missing.');
  if (!plasterEnd) missingSections.push('Plastering total in "m2" is missing.');
  if (!steelStart) missingSections.push('Section "plinth beam shuttering" is missing.');
  if (!steelEnd) missingSections.push('Steel total in "Tons" is missing.');
  if (!earthStart) missingSections.push('Earth work quantity in "sft" is missing.');
  if (missingSections.length) {
    throw new WorkbookFormatError("The measurement sections could not be identified.", missingSections);
  }

  const block = collectMeasure(wb, blockStart, blockEnd, {
    title: "Annexure - Measurement Sheet : Block Work (incl. plinth beam concreting)",
  });
  const plaster = collectMeasure(wb, plasterStart, plasterEnd, {
    title: "Annexure - Measurement Sheet : Plastering",
  });
  const steel = collectMeasure(wb, steelStart, steelEnd, {
    title: "Annexure - Measurement Sheet : Shuttering & Steel",
    cols: 9,
  });
  const earth = collectMeasure(wb, earthStart, wb.maxRow, {
    title: "Annexure - Measurement Sheet : Earth Work & PCC",
    cols: 9,
  });

  return { items, taxable, cgst, sgst, grand, block, plaster, steel, earth };
}

function normalized(value) {
  return String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

function findRow(wb, from, to, predicate) {
  for (let r = Math.max(1, from); r <= Math.min(to, wb.maxRow); r++) {
    if (predicate(r)) return r;
  }
  return null;
}

function findLastRow(wb, from, to, predicate) {
  for (let r = Math.min(to, wb.maxRow); r >= Math.max(1, from); r--) {
    if (predicate(r)) return r;
  }
  return null;
}

function rowContains(wb, row, values) {
  const wanted = values.map(normalized);
  return rowVals(wb, row, 12).map(normalized).some((value) => wanted.includes(value));
}

function hasAnyValue(wb, row, maxCol = 12) {
  return !rowEmpty(rowVals(wb, row, maxCol));
}

function previousMeaningfulRow(wb, before) {
  for (let r = before - 1; r >= 1; r--) {
    if (hasAnyValue(wb, r)) return r;
  }
  return null;
}

function findImmediateHeader(wb, dataRow) {
  const prior = previousMeaningfulRow(wb, dataRow);
  if (!prior) return null;
  return normalized(wb.cell(prior, 2)).includes("particular") && normalized(wb.cell(prior, 3)) === "nos"
    ? prior
    : null;
}

function findImmediateColumnHeader(wb, dataRow) {
  const prior = previousMeaningfulRow(wb, dataRow);
  if (!prior) return null;
  const text = rowVals(wb, prior, 9).map(normalized).join(" ");
  return /column 1|particular|length/.test(text) ? prior : null;
}

function isHeaderLike(desc) {
  const d = String(desc || "").trim();
  if (!d) return false;
  if (/^PLASTERING$/i.test(d)) return true;
  if (/inner side|outer side|Security & office|Block masonry|plinth beam|BUILDING|Total steel|Total block|Earth work/i.test(d))
    return true;
  return false;
}

function collectMeasure(wb, from, to, { title, cols = 8 }) {
  const rows = [];
  for (let r = from; r <= to; r++) {
    const vals = rowVals(wb, r, cols);
    if (rowEmpty(vals)) {
      if (rows.length && rows[rows.length - 1].kind !== "spacer") rows.push({ kind: "spacer" });
      continue;
    }
    const a = vals[0];
    const b = vals[1];
    const headerText = String(b || a || "").trim();
    const hasDims = vals.slice(2).some((v) => v !== "" && v != null);
    if (headerText && !hasDims && isHeaderLike(headerText)) {
      rows.push({ kind: "section", text: headerText });
      continue;
    }
    if (String(b).toUpperCase().includes("PARTICULAR") || String(a).toUpperCase() === "SL NO") {
      rows.push({
        kind: "header",
        cells: ["SL NO", "PARTICULAR", "NOS", "LENGTH", "BREADTH", "HEIGHT", "QUANTITY", "REMARKS"].slice(0, cols),
      });
      continue;
    }
    rows.push({
      kind: "data",
      cells: vals.map((v) => (typeof v === "number" ? v : String(v || "").trim())),
      bold: typeof b === "string" && /total/i.test(b),
    });
  }
  while (rows.length && rows[rows.length - 1].kind === "spacer") rows.pop();
  fillBlankSerials(rows);
  return { title, rows, cols };
}

function fillBlankSerials(rows) {
  let serial = 0;
  for (const row of rows) {
    if (row.kind === "section" || row.kind === "header") {
      serial = 0;
      continue;
    }
    if (row.kind !== "data") continue;
    const particular = String(row.cells[1] ?? "").trim();
    if (!particular || /^total/i.test(particular)) continue;
    const existing = row.cells[0];
    if (existing !== "" && existing != null) {
      const parsed = Number(existing);
      if (Number.isFinite(parsed)) serial = parsed;
      continue;
    }
    serial += 1;
    row.cells[0] = serial;
  }
}

function invoicePage(model) {
  const { seller, billTo, bank } = meta;
  const itemRows = model.items
    .map(
      (it) => `<tr>
      <td class="c">${it.sl}</td>
      <td>${esc(it.desc)}</td>
      <td>${esc(it.unit)}</td>
      <td class="r">${qty(it.qty)}</td>
      <td class="r">${plain(it.rate)}</td>
      <td></td>
      <td class="r nar">${money(it.amount)}</td>
    </tr>`
    )
    .join("");
  const pct = `${meta.gstRate * 100}%`;
  return `<section class="page">
    <h1>${esc(meta.title)}</h1>
    <table class="box invoice">
      <colgroup>
        <col style="width:6%"><col style="width:36%"><col style="width:8%">
        <col style="width:12%"><col style="width:10%"><col style="width:10%"><col style="width:18%">
      </colgroup>
      <tr>
        <td colspan="3" class="seller">
          <div class="b">${esc(seller.name)}</div>
          <div>Address: ${esc(seller.address1)}</div>
          <div class="indent">${esc(seller.address2)}</div>
          <div class="mt">GSTIN ${esc(seller.gstin)} &nbsp;&nbsp; Mobile: ${esc(seller.mobile)}</div>
          <div>Email ${esc(seller.email)}</div>
        </td>
        <td class="c headcell">Invoice No</td>
        <td class="c headcell">Invoice Date</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td colspan="3"></td>
        <td class="c">${esc(meta.invoiceNo)}</td>
        <td class="c">${esc(meta.invoiceDate)}</td>
        <td></td>
        <td></td>
      </tr>
      <tr class="gap"><td colspan="7"></td></tr>
      <tr>
        <td colspan="7"><span class="b">BILL TO</span></td>
      </tr>
      <tr>
        <td colspan="7" class="b">${esc(billTo.name)}</td>
      </tr>
      <tr>
        <td colspan="7">Address: ${esc(billTo.address1)}</td>
      </tr>
      <tr>
        <td colspan="7">${esc(billTo.address2)}</td>
      </tr>
      <tr>
        <td colspan="5">GSTIN: ${esc(billTo.gstin)}</td>
        <td colspan="2">Place of Supply : ${esc(billTo.placeOfSupply)}</td>
      </tr>
      <tr>
        <td colspan="5">Mobile : ${esc(billTo.mobile)}</td>
        <td colspan="2">PAN Number: ${esc(billTo.pan)}</td>
      </tr>
      <tr class="peach">
        <th>Sl No.</th><th>Items</th><th>Unit</th><th>Quantity</th><th>Price</th><th></th><th class="nar">Total</th>
      </tr>
      ${itemRows}
      <tr class="total-line">
        <td></td><td class="b">Total</td><td></td><td></td><td></td><td></td>
        <td class="r nar b">${money(model.taxable)}</td>
      </tr>
      <tr>
        <td></td><td class="r">CGST @${pct} -</td><td class="c">-</td><td class="c">-</td><td></td><td></td>
        <td class="r nar">${money(model.cgst)}</td>
      </tr>
      <tr>
        <td></td><td class="r">SGST @${pct} -</td><td class="c">-</td><td class="c">-</td><td></td><td></td>
        <td class="r nar">${money(model.sgst)}</td>
      </tr>
      <tr class="peach grand">
        <td colspan="6" class="c b">Total</td>
        <td class="r nar b">${money(model.grand)}</td>
      </tr>
      <tr class="peach">
        <th rowspan="2">HSN/SAC</th>
        <th rowspan="2">Taxable Value</th>
        <th colspan="2">CGST</th>
        <th colspan="2">SGST</th>
        <th rowspan="2" class="nar">Total Tax Amount</th>
      </tr>
      <tr class="peach">
        <th>Rate</th><th>Amount</th><th>Rate</th><th>Amount</th>
      </tr>
      <tr>
        <td></td>
        <td class="r">${money(model.taxable)}</td>
        <td class="r">${pct}</td>
        <td class="r">${money(model.cgst)}</td>
        <td class="r">${pct}</td>
        <td class="r">${money(model.sgst)}</td>
        <td class="r nar">${money(r2(model.cgst + model.sgst))}</td>
      </tr>
      <tr class="b">
        <td>Total</td>
        <td class="r">${money(model.taxable)}</td>
        <td></td>
        <td class="r">${money(model.cgst)}</td>
        <td></td>
        <td class="r">${money(model.sgst)}</td>
        <td class="r nar">${money(r2(model.cgst + model.sgst))}</td>
      </tr>
      <tr class="sign">
        <td colspan="2" class="bank">
          <div class="b">Bank Details :</div>
          <div>Bank: &nbsp; ${esc(bank.bank)}</div>
          <div>Name: &nbsp; ${esc(bank.name)}</div>
          <div>IFSC Code: ${esc(bank.ifsc)}</div>
          <div>A/c No: &nbsp; ${esc(bank.account)}</div>
          <div>Branch : &nbsp; ${esc(bank.branch)}</div>
        </td>
        <td colspan="2" class="c signbox">
          <div>Engineer Signature</div>
          <div>for approval</div>
        </td>
        <td colspan="2" class="c signbox">
          <div class="mt2">Authorised Signatory</div>
          <div>${esc(meta.signatory)}</div>
        </td>
        <td class="c signbox nar">
          <div class="mt2">Receiver's</div>
          <div>Signature</div>
        </td>
      </tr>
    </table>
  </section>`;
}

function subLine() {
  return `Annexure to ${meta.title} No. ${meta.invoiceNo} dated ${meta.invoiceDate} | ${meta.workName}`;
}

function measurePages(sheet) {
  const colCount = sheet.cols || 8;
  const hdr = ["SL NO", "PARTICULAR", "NOS", "LENGTH", "BREADTH", "HEIGHT", "QUANTITY", "REMARKS", "UNIT"];
  const chunks = paginateMeasureRows(sheet.rows, 58);
  const widths = colCount === 9
    ? [5, 29, 8, 10, 10, 10, 10, 10, 8]
    : [6, 34, 7, 10, 10, 10, 11, 12];
  const colgroup = `<colgroup>${widths
    .slice(0, colCount)
    .map((width) => `<col style="width:${width}%">`)
    .join("")}</colgroup>`;

  return chunks
    .map((rows) => {
      const lines = [];
      for (const row of rows) {
    if (row.kind === "spacer") {
      lines.push(`<tr class="spacer"><td colspan="${colCount}"></td></tr>`);
      continue;
    }
    if (row.kind === "section") {
      lines.push(`<tr class="section"><td colspan="${colCount}">${esc(row.text)}</td></tr>`);
      continue;
    }
    if (row.kind === "header") {
      lines.push(
        `<tr class="peach${row.repeated ? " repeat-header" : ""}">${hdr
          .slice(0, colCount)
          .map((h) => `<th>${h}</th>`)
          .join("")}</tr>`
      );
      continue;
    }
    const cells = row.cells.slice(0, colCount);
    while (cells.length < colCount) cells.push("");
    const cls = row.bold ? " class=\"b\"" : "";
    lines.push(
      `<tr${cls}>${cells
        .map((v, i) => {
          const align = i === 1 || i === 7 ? "" : " class=\"r\"";
          const show = typeof v === "number" ? (i === 6 ? qty(v, 2) : prettyNum(v)) : esc(v);
          return `<td${i === 1 ? "" : align}>${show}</td>`;
        })
        .join("")}</tr>`
    );
      }
      return `<section class="page measure-page">
    <h2>${esc(sheet.title)}</h2>
    <p class="sub">${esc(subLine())}</p>
    <table class="grid compact measure">
      ${colgroup}
      ${lines.join("\n")}
    </table>
  </section>`;
    })
    .join("\n");
}

function paginateMeasureRows(rows, maxRows) {
  const chunks = [];
  let current = [];
  let used = 0;
  let headerRow = null;

  for (const row of rows) {
    if (row.kind === "header") headerRow = row;
    const weight = row.kind === "spacer" ? 0.5 : row.kind === "section" ? 1.5 : 1;

    if (current.length && used + weight > maxRows) {
      let carry = null;
      if (current[current.length - 1].kind === "section") {
        carry = current.pop();
      }
      chunks.push(current);
      current = headerRow ? [{ ...headerRow, repeated: true }] : [];
      used = current.length;
      if (carry) {
        current.push(carry);
        used += 1.5;
      }
    }

    current.push(row);
    used += weight;
  }

  if (current.length) chunks.push(current);
  return chunks;
}

function prettyNum(v) {
  if (typeof v !== "number") return esc(v);
  if (Number.isInteger(v)) return String(v);
  const s = v.toFixed(4).replace(/0+$/, "").replace(/\.$/, "");
  return s;
}

function css(bg) {
  return `* { box-sizing: border-box; }
@page { size: A4; margin: 0; }
html, body { margin: 0; padding: 0; background: #d9d9d9; }
body { font-family: Calibri, Arial, sans-serif; font-size: 10px; color: #111; }
.toolbar { position: sticky; top: 0; z-index: 5; background: #222; color: #fff; padding: 10px 16px; display: flex; gap: 12px; align-items: center; }
.toolbar button { font: 14px Calibri, sans-serif; padding: 6px 14px; cursor: pointer; }
.sheet { width: 210mm; margin: 12px auto; }
.page {
  width: 210mm; height: 297mm; background: #fff url("${bg}") center / 100% 100% no-repeat;
  padding: 34mm 11mm 20mm 11mm; page-break-after: always; position: relative;
  overflow: hidden;
}
h1 { font-size: 18px; margin: 0 0 6px; letter-spacing: 0.3px; }
h2 { font-size: 14px; margin: 0 0 2px; }
.sub { margin: 0 0 8px; font-style: italic; font-size: 8.5px; color: #333; }
.note { margin-top: 10px; font-style: italic; font-size: 8px; }
table { border-collapse: collapse; width: 100%; }
.box td, .box th, .grid td, .grid th { border: 0.6pt solid #000; padding: 2px 4px; vertical-align: middle; }
.invoice { table-layout: fixed; }
.invoice td, .invoice th { overflow-wrap: anywhere; }
.peach th, .peach td, tr.peach td, tr.peach th, .grand td { background: #FCE4D6; }
.b { font-weight: 700; }
.c { text-align: center; }
.r { text-align: right; }
.nar { font-family: "Arial Narrow", Calibri, Arial, sans-serif; }
.seller .indent { padding-left: 52px; }
.mt { margin-top: 4px; }
.mt2 { margin-top: 28px; }
.headcell { font-weight: 700; }
tr.gap td { border-left: none; border-right: none; height: 6px; }
.total-line td { border-top: 1.4pt solid #000; }
.sign td { height: 78px; vertical-align: top; }
.signbox { vertical-align: bottom !important; padding-bottom: 8px; }
.bank div { line-height: 1.45; }
.compact td, .compact th { font-size: 8px; padding: 1.5px 3px; }
.measure tr.section td { background: #FCE4D6; font-weight: 700; }
.measure tr.spacer td { border: none; height: 6px; }
.measure { table-layout: fixed; }
.measure td, .measure th { overflow-wrap: anywhere; }
@media print {
  body { background: none; }
  .toolbar { display: none; }
  .sheet { margin: 0; }
  .page { margin: 0; box-shadow: none; overflow: hidden; }
}`;
}

function htmlDoc(model) {
  const bg = letterheadDataUri();
  return `<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<title>${esc(meta.title)} ${esc(meta.invoiceNo)}</title>
<style>${css(bg)}</style>
</head>
<body>
<div class="toolbar">
  <strong>Kshma billing</strong>
  <span>Invoice ${esc(meta.invoiceNo)} — print on A4, background graphics on</span>
  <button onclick="window.print()">Print / Save PDF</button>
</div>
<div class="sheet">
${invoicePage(model)}
${measurePages(model.block)}
${measurePages(model.plaster)}
${measurePages(model.steel)}
${measurePages(model.earth)}
</div>
</body>
</html>`;
}

function createInvoiceFromWorkbook(workbookPath) {
  if (!fs.existsSync(workbookPath)) {
    throw new Error(`Workbook not found: ${workbookPath}`);
  }
  const wb = loadWorkbook(workbookPath);
  const model = buildModel(wb);
  return { model, html: htmlDoc(model) };
}

function main() {
  const xlsxArg = process.argv[2];
  const xlsxPath = xlsxArg
    ? path.resolve(xlsxArg)
    : path.join(ROOT, "source", "HARSHA BILL 2.xlsx");
  if (!fs.existsSync(xlsxPath)) {
    console.error("Workbook not found:", xlsxPath);
    process.exit(1);
  }
  const { model, html } = createInvoiceFromWorkbook(xlsxPath);
  const outDir = path.join(ROOT, "output");
  fs.mkdirSync(outDir, { recursive: true });
  const outFile = path.join(outDir, `Proforma_Invoice_${meta.invoiceNo}.html`);
  fs.writeFileSync(outFile, html);
  console.log("Wrote", outFile);
  console.log("Taxable", money(model.taxable), "Grand", money(model.grand));
}

if (require.main === module) main();

module.exports = {
  WorkbookFormatError,
  buildModel,
  createInvoiceFromWorkbook,
  htmlDoc,
  money,
};
