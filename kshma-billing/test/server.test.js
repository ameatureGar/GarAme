const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");

const workbook = path.join(__dirname, "..", "source", "HARSHA BILL 2.xlsx");

test("upload endpoint returns a generated invoice", async (t) => {
  const { createBillingServer } = require("../scripts/server");
  const app = createBillingServer();
  await new Promise((resolve) => app.server.listen(0, "127.0.0.1", resolve));
  t.after(() => new Promise((resolve) => app.server.close(resolve)));

  const address = app.server.address();
  const response = await fetch(`http://127.0.0.1:${address.port}/api/generate`, {
    method: "POST",
    headers: {
      "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "x-filename": "HARSHA BILL 2.xlsx",
    },
    body: fs.readFileSync(workbook),
  });

  assert.equal(response.status, 200);
  const result = await response.json();
  assert.equal(result.summary.billItems, 9);
  assert.equal(result.summary.measurementRows, 366);
  assert.equal(result.summary.grand, 627714.34);

  const invoice = await fetch(`http://127.0.0.1:${address.port}${result.invoiceUrl}`);
  assert.equal(invoice.status, 200);
  assert.match(await invoice.text(), /PROFORMA INVOICE/);
});

test("upload endpoint rejects non-XLSX input", async (t) => {
  const { createBillingServer } = require("../scripts/server");
  const app = createBillingServer();
  await new Promise((resolve) => app.server.listen(0, "127.0.0.1", resolve));
  t.after(() => new Promise((resolve) => app.server.close(resolve)));

  const address = app.server.address();
  const response = await fetch(`http://127.0.0.1:${address.port}/api/generate`, {
    method: "POST",
    headers: { "content-type": "text/plain", "x-filename": "notes.txt" },
    body: "not a workbook",
  });

  assert.equal(response.status, 415);
  const result = await response.json();
  assert.match(result.error, /XLSX/);
});
