const express      = require('express');
const ExcelJS      = require('exceljs');
const htmlPdfNode  = require('html-pdf-node');
const cors         = require('cors');

const app  = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '50mb' }));
app.use(cors({ origin: '*', methods: ['POST', 'OPTIONS'] }));

app.post('/export-pdf', async (req, res) => {
  try {
    const payload = req.body;
    console.log('Received request, payload keys:', Object.keys(payload));

    // ── Extract data from ThoughtSpot payload ────────────────────────────
    const embedData = payload?.payload?.data?.embedAnswerData
                   ?? payload?.data?.embedAnswerData
                   ?? payload?.embedAnswerData
                   ?? payload?.reportBookData
                   ?? payload;

    // Try multiple known payload shapes ThoughtSpot uses
    let columnHeaders = [];
    let rows = [];

    // Shape 1: embedAnswerData with columns + columnDataLite
    if (embedData?.columns && embedData?.data) {
      columnHeaders = embedData.columns.map(col => col?.column?.name ?? 'Column');
      const columnDataLite = (embedData.data?.[0] ?? embedData.data)?.columnDataLite ?? [];
      const numRows = columnDataLite[0]?.dataValue?.length ?? 0;
      for (let r = 0; r < numRows; r++) {
        rows.push(columnDataLite.map(col => col.dataValue[r] ?? ''));
      }
    }
    // Shape 2: reportBookData structure
    else if (payload?.reportBookData) {
      const reportBook = payload.reportBookData;
      const vizData = Object.values(reportBook)?.[0]?.vizData;
      const firstViz = vizData ? Object.values(vizData)?.[0] : null;
      const dataSet = firstViz?.dataSets?.PINBOARD_VIZ ?? firstViz?.dataSets?.SEARCH_BAR_VIZ;
      if (dataSet) {
        columnHeaders = (dataSet.columns ?? []).map(c => c?.column?.name ?? 'Column');
        const columnDataLite = (Array.isArray(dataSet.data)
          ? dataSet.data[0]
          : dataSet.data)?.columnDataLite ?? [];
        const numRows = columnDataLite[0]?.dataValue?.length ?? 0;
        for (let r = 0; r < numRows; r++) {
          rows.push(columnDataLite.map(col => col.dataValue[r] ?? ''));
        }
      }
    }

    if (columnHeaders.length === 0) {
      console.error('Could not extract columns. embedData:', JSON.stringify(embedData).slice(0, 300));
      return res.status(400).json({ error: 'Could not extract column data from payload' });
    }

    console.log(`Processing ${rows.length} rows x ${columnHeaders.length} columns`);

    // ── Build HTML table ─────────────────────────────────────────────────
    const html = buildPrintableHTML(columnHeaders, rows);

    // ── Convert HTML → PDF using html-pdf-node ───────────────────────────
    const options = {
      format: 'A4',
      landscape: false,
      printBackground: true,
      margin: { top: '15mm', bottom: '15mm', left: '12mm', right: '12mm' },
    };

    const file = { content: html };
    const pdfBuffer = await htmlPdfNode.generatePdf(file, options);

    console.log('PDF generated successfully');

res.set({
  'Content-Type': 'application/pdf',
  'Content-Disposition': 'attachment; filename="thoughtspot-export.pdf"',
  'Content-Length': pdfBuffer.length,
  'Access-Control-Expose-Headers': 'Content-Disposition' // Important for CORS
});

return res.status(200).send(pdfBuffer);

    res.send(pdfBuffer);

  } catch (err) {
    console.error('Export failed:', err.message);
    res.status(500).json({ error: 'PDF generation failed', detail: err.message });
  }
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

app.listen(PORT, () => {
  console.log(`ts-pdf-exporter running on port ${PORT}`);
});

// ── Helper: build print-ready HTML ───────────────────────────────────────────
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
    padding: 0;
  }
  .header {
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
    padding: 12px 0 8px 0;
    margin-bottom: 10px;
    border-bottom: 2px solid #1e3a5f;
  }
  .title { font-size: 13pt; font-weight: 700; color: #1e3a5f; }
  .meta  { font-size: 7.5pt; color: #6b7280; text-align: right; }
  table  { width: 100%; border-collapse: collapse; }
  thead  { display: table-header-group; }
  thead tr {
    background: #1e3a5f;
    color: #fff;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
  th {
    padding: 6px 8px;
    text-align: left;
    font-weight: 600;
    font-size: 8.5pt;
    border: 1px solid #1e3a5f;
    white-space: nowrap;
  }
  td {
    padding: 4px 8px;
    border: 1px solid #e2e8f0;
    border-top: none;
    vertical-align: top;
  }
  tr.even {
    background: #f0f4ff;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
  tr.odd  { background: #ffffff; }
  .footer {
    margin-top: 8px;
    font-size: 7pt;
    color: #9ca3af;
    text-align: center;
  }
</style>
</head>
<body>
  <div class="header">
    <div class="title">ThoughtSpot Export</div>
    <div class="meta">
      Generated: ${new Date().toLocaleString()}<br/>
      ${rows.length} rows &bull; ${headers.length} columns
    </div>
  </div>
  <table>${thead}${tbody}</table>
  <div class="footer">Exported from ThoughtSpot</div>
</body>
</html>`;
}