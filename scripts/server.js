const fs = require("fs");
const http = require("http");
const os = require("os");
const path = require("path");
const { randomUUID } = require("crypto");
const { WorkbookFormatError, createInvoiceFromWorkbook, meta, resolveInvoiceNo } = require("./generate-invoice");

const ROOT = path.join(__dirname, "..");
const PUBLIC = path.join(ROOT, "public");
const MAX_UPLOAD_BYTES = 20 * 1024 * 1024;

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
};

function json(res, status, value) {
  const body = Buffer.from(JSON.stringify(value));
  res.writeHead(status, {
    "content-type": "application/json; charset=utf-8",
    "content-length": body.length,
    "cache-control": "no-store",
  });
  res.end(body);
}

function readUpload(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let total = 0;
    req.on("data", (chunk) => {
      total += chunk.length;
      if (total > MAX_UPLOAD_BYTES) {
        reject(Object.assign(new Error("The workbook is larger than 20 MB."), { status: 413 }));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function readHeader(req, name) {
  const raw = req.headers[name];
  if (raw == null || raw === "") return undefined;
  try {
    return decodeURIComponent(String(raw));
  } catch {
    throw Object.assign(
      new Error("Enter an invoice number using letters, numbers, spaces, or - / _ ."),
      { status: 400 }
    );
  }
}

function measurementRowCount(model) {
  return [model.block, model.plaster, model.steel, model.earth]
    .flatMap((section) => section.rows).length;
}

function createBillingServer() {
  const invoices = new Map();

  const server = http.createServer(async (req, res) => {
    const url = new URL(req.url, "http://localhost");

    if (req.method === "GET" && url.pathname === "/api/defaults") {
      return json(res, 200, { invoiceNo: meta.invoiceNo });
    }

    if (req.method === "POST" && url.pathname === "/api/generate") {
      const filename = decodeURIComponent(String(req.headers["x-filename"] || ""));
      const contentType = String(req.headers["content-type"] || "").toLowerCase();
      if (!filename.toLowerCase().endsWith(".xlsx") || !contentType.includes("spreadsheetml")) {
        return json(res, 415, { error: "Please upload an XLSX workbook." });
      }

      let invoiceNo;
      try {
        invoiceNo = resolveInvoiceNo(readHeader(req, "x-invoice-no"));
      } catch (error) {
        return json(res, error.status || 400, { error: error.message });
      }

      let tempFile;
      try {
        const data = await readUpload(req);
        if (data.length < 4 || data[0] !== 0x50 || data[1] !== 0x4b) {
          return json(res, 400, { error: "The uploaded file is not a valid XLSX workbook." });
        }
        tempFile = path.join(os.tmpdir(), `kshma-upload-${randomUUID()}.xlsx`);
        fs.writeFileSync(tempFile, data);
        const result = createInvoiceFromWorkbook(tempFile, { invoiceNo });
        const id = randomUUID();
        invoices.set(id, { html: result.html, createdAt: Date.now() });
        while (invoices.size > 20) invoices.delete(invoices.keys().next().value);

        return json(res, 200, {
          id,
          invoiceUrl: `/invoice/${id}`,
          printUrl: `/invoice/${id}?print=1`,
          summary: {
            filename,
            invoiceNo: result.invoiceNo,
            billItems: result.model.items.length,
            measurementRows: measurementRowCount(result.model),
            taxable: result.model.taxable,
            grand: result.model.grand,
          },
        });
      } catch (error) {
        if (error instanceof WorkbookFormatError) {
          return json(res, 422, { error: error.message, details: error.details });
        }
        const status = error.status || 400;
        return json(res, status, { error: error.message || "The workbook could not be processed." });
      } finally {
        if (tempFile && fs.existsSync(tempFile)) fs.rmSync(tempFile, { force: true });
      }
    }

    if (req.method === "GET" && url.pathname.startsWith("/invoice/")) {
      const id = url.pathname.slice("/invoice/".length);
      const invoice = invoices.get(id);
      if (!invoice) return json(res, 404, { error: "This generated invoice has expired." });
      const printScript = url.searchParams.get("print") === "1"
        ? '<script>addEventListener("load",()=>setTimeout(()=>window.print(),250));</script>'
        : "";
      const html = invoice.html.replace("</body>", `${printScript}</body>`);
      res.writeHead(200, { "content-type": TYPES[".html"], "cache-control": "no-store" });
      return res.end(html);
    }

    if (req.method === "GET") {
      const requested = url.pathname === "/" ? "index.html" : url.pathname.slice(1);
      const allowed = new Set(["index.html", "styles.css", "app.js"]);
      if (allowed.has(requested)) {
        const file = path.join(PUBLIC, requested);
        const body = fs.readFileSync(file);
        res.writeHead(200, {
          "content-type": TYPES[path.extname(file)] || "application/octet-stream",
          "content-length": body.length,
        });
        return res.end(body);
      }
    }

    json(res, 404, { error: "Not found." });
  });

  return { server, invoices };
}

if (require.main === module) {
  const port = Number(process.env.PORT || 4173);
  const { server } = createBillingServer();
  server.listen(port, "127.0.0.1", () => {
    console.log(`Kshma Billing is running at http://127.0.0.1:${port}`);
    console.log("Press Ctrl+C to stop it.");
  });
}

module.exports = { createBillingServer };

