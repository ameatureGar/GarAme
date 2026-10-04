# Kshma billing

Generates a **proforma invoice** that matches the Kshma Group letterhead, then appends annexure pages from the measurement workbook.

## What you get

1. Proforma invoice (same layout as the sample bill)
2. Bill abstract — Total / Present / Balance quantities
3. Measurement sheets — block work, plastering, shuttering & steel, earth work & PCC

Every page uses the Word letterhead as a full-page background.

## Upload a workbook in the browser

1. From this folder run:

```bat
npm start
```

2. Open `http://127.0.0.1:4173` in a browser.
3. Choose or drag in an `.xlsx` measurement workbook.
4. Review the detected item count and totals, then use **View invoice** or **Print / Save PDF**.

The upload stays on this computer. It is written to a temporary file only while the workbook is being read, then deleted. Generated invoices remain available until the local server stops. The importer finds the bill and annexure sections from their headings, so it supports both the original workbook and the formatted version with inserted rows.

## Generate from the command line

1. Put the latest workbook in `source/HARSHA BILL 2.xlsx` (or pass another path).
2. Edit bill header, GSTIN, bank and invoice number in `data/invoice-meta.json`.
3. From this folder run:

```bat
node scripts/generate-invoice.js
```

Or with another workbook:

```bat
node scripts/generate-invoice.js "C:\path\to\workbook.xlsx"
```

4. Open `output/Proforma_Invoice_<number>.html` in the browser.
5. Print → Save as PDF. Use **A4**, **background graphics on**, margins **none**.

Present QTS (this bill) is what appears on the invoice. Totals and GST (9% CGST + 9% SGST) are recalculated from those quantities after rounding to 2 decimals, matching the sample proforma.

## Checks

Run:

```bat
npm test
```

The checks cover the original workbook, the formatted workbook when it is present in Downloads, the upload endpoint, invoice retrieval, totals, pagination, and invalid file rejection.
