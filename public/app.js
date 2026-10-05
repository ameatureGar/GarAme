const invoiceNoInput = document.querySelector("#invoice-no");
const fileInput = document.querySelector("#file-input");
const chooseButton = document.querySelector("#choose-button");
const dropZone = document.querySelector("#drop-zone");
const progress = document.querySelector("#progress");
const summary = document.querySelector("#summary");
const result = document.querySelector("#result");
const errorPanel = document.querySelector("#error");

const money = new Intl.NumberFormat("en-IN", {
  style: "currency",
  currency: "INR",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

let invoiceNoEdited = false;
invoiceNoInput.addEventListener("input", () => {
  invoiceNoEdited = true;
});
loadDefaultInvoiceNo();

chooseButton.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", () => {
  if (fileInput.files[0]) processFile(fileInput.files[0]);
});

for (const event of ["dragenter", "dragover"]) {
  dropZone.addEventListener(event, (e) => {
    e.preventDefault();
    dropZone.classList.add("dragging");
  });
}
for (const event of ["dragleave", "drop"]) {
  dropZone.addEventListener(event, (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragging");
  });
}
dropZone.addEventListener("drop", (e) => {
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
});

async function processFile(file) {
  resetPanels();
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    return showError("Please upload an XLSX workbook.");
  }
  if (file.size > 20 * 1024 * 1024) {
    return showError("The workbook is larger than 20 MB.");
  }
  const invoiceNo = invoiceNoInput.value.trim();
  if (!invoiceNo) {
    invoiceNoInput.focus();
    return showError("Enter an invoice number.");
  }

  progress.hidden = false;
  document.querySelector("#progress-name").textContent = file.name;
  chooseButton.disabled = true;
  invoiceNoInput.disabled = true;
  chooseButton.textContent = "Processing…";

  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      headers: {
        "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "x-filename": encodeURIComponent(file.name),
        "x-invoice-no": encodeURIComponent(invoiceNo),
      },
      body: file,
    });
    const payload = await response.json();
    if (!response.ok) throw Object.assign(new Error(payload.error), { details: payload.details });
    showResult(payload);
  } catch (error) {
    showError(error.message || "The workbook could not be processed.", error.details || []);
  } finally {
    progress.hidden = true;
    chooseButton.disabled = false;
    invoiceNoInput.disabled = false;
    chooseButton.textContent = "Choose another Excel file";
  }
}

async function loadDefaultInvoiceNo() {
  try {
    const response = await fetch("/api/defaults");
    if (!response.ok) return;
    const data = await response.json();
    if (data.invoiceNo && !invoiceNoEdited) invoiceNoInput.value = data.invoiceNo;
  } catch {
    // The field already shows the saved default.
  }
}

function resetPanels() {
  progress.hidden = true;
  summary.hidden = true;
  result.hidden = true;
  errorPanel.hidden = true;
  document.querySelector("#error-details").replaceChildren();
}

function showResult(payload) {
  const data = payload.summary;
  document.querySelector("#file-name").textContent = data.filename;
  document.querySelector("#used-invoice-no").textContent = data.invoiceNo;
  document.querySelector("#bill-items").textContent = data.billItems;
  document.querySelector("#measurement-rows").textContent = data.measurementRows;
  document.querySelector("#taxable").textContent = money.format(data.taxable);
  document.querySelector("#grand").textContent = money.format(data.grand);
  document.querySelector("#view-link").href = payload.invoiceUrl;
  document.querySelector("#print-link").href = payload.printUrl;
  summary.hidden = false;
  result.hidden = false;
  summary.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function showError(message, details = []) {
  document.querySelector("#error-message").textContent = message;
  const list = document.querySelector("#error-details");
  list.replaceChildren(...details.map((detail) => {
    const item = document.createElement("li");
    item.textContent = detail;
    return item;
  }));
  errorPanel.hidden = false;
  errorPanel.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

