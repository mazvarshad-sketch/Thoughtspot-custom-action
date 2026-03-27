const express    = require('express');
const ExcelJS    = require('exceljs');
const puppeteer  = require('puppeteer');
const cors       = require('cors');
const path       = require('path');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Middleware ────────────────────────────────────────────────────────────────

app.use(express.json({ limit: '50mb' }));  // ThoughtSpot payloads can be large

// IMPORTANT: replace the origin below with your exact ThoughtSpot instance URL
app.use(cors({
  origin: '*',
  methods: ['POST', 'OPTIONS'],
}));


// ── Main endpoint — receives ThoughtSpot URL action POST ──────────────────────
//
// ThoughtSpot URL actions send a POST request with this body shape:
// {
//   "actionId": "export-to-pdf",
//   "payload": {
//     "data": {
//       "embedAnswerData": {
//         "columns": [{ "column": { "name": "Region" } }, ...],
//         "data": [{ "columnDataLite": [{ "dataValue": ["North","South"] }, ...] }]
//       }
//     }
//   }
// }

app.post('/export-pdf', async (req, res) => {
  try {
    console.log('Received export request from ThoughtSpot');

    // ── 1. Extract data from ThoughtSpot payload ──────────────────────────
    const payload      = req.body;
    const embedData    = payload?.payload?.data?.embedAnswerData
                      ?? payload?.data?.embedAnswerData
                      ?? payload?.embedAnswerData;

    if (!embedData) {
      return res.status(400).json({ error: 'No embedAnswerData found in payload' });
    }

    const columns        = embedData.columns ?? [];
    const columnHeaders  = columns.map(col => col?.column?.name ?? 'Column');
    const columnDataLite = (embedData.data?.[0] ?? embedData.data)?.columnDataLite ?? [];
    const numRows        = columnDataLite[0]?.dataValue?.length ?? 0;

    // Convert column-major data to row-major (array of row arrays)
    const rows = [];
    for (let r = 0; r < numRows; r++) {
      rows.push(columnDataLite.map(col => col.dataValue[r] ?? ''));
    }

    console.log(`Processing ${numRows} rows × ${columnHeaders.length} columns`);

    // ── 2. Build XLSX in memory using ExcelJS ─────────────────────────────
    const workbook  = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('ThoughtSpot Export');

    // Style the header row
    worksheet.addRow(columnHeaders);
    const headerRow = worksheet.getRow(1);
    headerRow.eachCell(cell => {
      cell.font       = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      cell.fill       = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } };
      cell.alignment  = { vertical: 'middle', horizontal: 'left' };
      cell.border     = { bottom: { style: 'thin', color: { argb: 'FF1E3A5F' } } };
    });
    headerRow.height = 22;

    // Add data rows with alternating row shading
    rows.forEach((row, i) => {
      const excelRow = worksheet.addRow(row);
      excelRow.height = 18;
      if (i % 2 === 0) {
        excelRow.eachCell(cell => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F4FF' } };
        });
      }
      excelRow.eachCell(cell => {
        cell.border = {
          bottom: { style: 'hair', color: { argb: 'FFE2E8F0' } },
        };
        cell.alignment = { vertical: 'middle' };
      });
    });

    // Auto-fit column widths (cap at 40 chars)
    worksheet.columns.forEach((col, i) => {
      const maxLen = Math.max(
        columnHeaders[i]?.length ?? 10,
        ...rows.map(r => String(r[i] ?? '').length)
      );
      col.width = Math.min(maxLen + 4, 40);
    });

    // Write workbook to buffer (we don't save to disk — all in memory)
    const xlsxBuffer = await workbook.xlsx.writeBuffer();


    // ── 3. Convert XLSX → HTML table → PDF via Puppeteer ─────────────────
    //
    // Strategy: render the data as a styled HTML table in headless Chrome,
    // then use Chrome's built-in PDF engine to produce A4 portrait PDF.
    // This gives the best, most reliable PDF output.

    const htmlContent = buildPrintableHTML(columnHeaders, rows);

    const browser = await puppeteer.launch({
      headless: 'new',
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',   // required on Railway/Docker
        '--disable-gpu',
      ],
    });

    const page = await browser.newPage();
    await page.setContent(htmlContent, { waitUntil: 'networkidle0' });

    const pdfBuffer = await page.pdf({
      format: 'A4',
      landscape: false,           // portrait as requested
      printBackground: true,      // needed to render coloured header/rows
      margin: {
        top:    '15mm',
        bottom: '15mm',
        left:   '12mm',
        right:  '12mm',
      },
    });

    await browser.close();
    console.log('PDF generated successfully');


    // ── 4. Stream the PDF back to the user's browser ──────────────────────
    //
    // ThoughtSpot opens the URL action response in the user's browser tab,
    // so setting Content-Disposition: attachment triggers a download.

    res.set({
      'Content-Type':        'application/pdf',
      'Content-Disposition': 'attachment; filename="thoughtspot-export.pdf"',
      'Content-Length':      pdfBuffer.length,
    });

    res.send(pdfBuffer);

  } catch (err) {
    console.error('Export failed:', err);
    res.status(500).json({ error: 'PDF generation failed', detail: err.message });
  }
});


// ── Health check endpoint (Railway uses this to verify the app is up) ─────────
app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});


// ── Start server ──────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`ts-pdf-exporter running on port ${PORT}`);
});


// ════════════════════════════════════════════════════════════════════════════
// HELPER — builds a clean print-ready HTML document from headers + rows
// ════════════════════════════════════════════════════════════════════════════

function buildPrintableHTML(headers, rows) {
  const escape = str =>
    String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');

  const thead = `<thead><tr>${
    headers.map(h => `<th>${escape(h)}</th>`).join('')
  }</tr></thead>`;

  const tbody = `<tbody>${
    rows.map((row, i) => `
      <tr class="${i % 2 === 0 ? 'even' : 'odd'}">
        ${headers.map((_, ci) => `<td>${escape(row[ci] ?? '')}</td>`).join('')}
      </tr>`
    ).join('')
  }</tbody>`;

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  @page { size: A4 portrait; margin: 0; }
  body {
    font-family: Arial, Helvetica, sans-serif;
    font-size: 9pt;
    color: #1a1a1a;
    background: #fff;
  }
  .header {
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
    padding: 0 0 8px 0;
    margin-bottom: 10px;
    border-bottom: 2px solid #1e3a5f;
  }
  .title   { font-size: 13pt; font-weight: 700; color: #1e3a5f; }
  .meta    { font-size: 7.5pt; color: #6b7280; text-align: right; }
  table    { width: 100%; border-collapse: collapse; }
  thead    { display: table-header-group; }
  thead tr { background: #1e3a5f; color: #fff; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  th       { padding: 6px 8px; text-align: left; font-weight: 600; font-size: 8.5pt; border: 1px solid #1e3a5f; white-space: nowrap; }
  td       { padding: 4px 8px; border: 1px solid #e2e8f0; border-top: none; vertical-align: top; }
  tr.even  { background: #f0f4ff; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  tr.odd   { background: #ffffff; }
  .footer  { margin-top: 8px; font-size: 7pt; color: #9ca3af; text-align: center; }
</style>
</head>
<body>
  <div class="header">
    <div class="title">ThoughtSpot Export</div>
    <div class="meta">Generated: ${new Date().toLocaleString()}<br/>${rows.length} rows &bull; ${headers.length} columns</div>
  </div>
  <table>${thead}${tbody}</table>
  <div class="footer">Exported from ThoughtSpot</div>
</body>
</html>`;
}
