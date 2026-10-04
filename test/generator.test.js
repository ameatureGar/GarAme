const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");

const ROOT = path.join(__dirname, "..");
const original = path.join(ROOT, "source", "HARSHA BILL 2.xlsx");
const formatted = path.join(
  process.env.USERPROFILE || "C:\\Users\\Raghav Chakravarthy",
  "Downloads",
  "Untitled spreadsheet.xlsx"
);

function verifyWorkbook(file, expectedMeasurementRows) {
  const { createInvoiceFromWorkbook } = require("../scripts/generate-invoice");
  const result = createInvoiceFromWorkbook(file);
  const measurementRows = [result.model.block, result.model.plaster, result.model.steel, result.model.earth]
    .flatMap((section) => section.rows).length;

  assert.equal(result.model.items.length, 9);
  assert.equal(measurementRows, expectedMeasurementRows);
  assert.equal(result.model.taxable, 531961.3);
  assert.equal(result.model.cgst, 47876.52);
  assert.equal(result.model.sgst, 47876.52);
  assert.equal(result.model.grand, 627714.34);
  assert.match(result.html, /PROFORMA INVOICE/);
  assert.match(result.html, /Concreting done by gang by mixture machine/);
  const door = result.model.block.rows.find((row) => row.kind === "data" && row.cells[1] === "Door");
  assert.equal(door.cells[0], 2);
  assert.doesNotMatch(result.html, /Bill Abstract/);
  assert.equal((result.html.match(/class="page/g) || []).length, 9);
}

test("generates the approved invoice from the original workbook", () => {
  verifyWorkbook(original, 366);
});

test("generates the same invoice from the formatted workbook", { skip: !fs.existsSync(formatted) }, () => {
  verifyWorkbook(formatted, 365);
});
