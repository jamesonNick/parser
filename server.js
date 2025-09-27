// server.js
const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');
const { v4: uuidv4 } = require('uuid');
const fs = require('fs');
const path = require('path');

const upload = multer({ dest: 'uploads/' });
const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static('public')); // serve frontend (index.html will go in /public)

const jobs = {}; // in-memory job status + results (for demo). For production use a DB/cache.

const headerCandidates = {
  POL: ['pol', 'port of loading', 'port loading'],
  POD: ['pod', 'port of discharge', 'port discharge'],
  RATES: ['rate', 'rates', 'freight', 'amount', 'price'],
  SIZE: ['size', 'container size', 'container-size'],
  CONTAINER_TYPE: ['container type', 'container', 'ctype', 'type', 'container-type'],
  CM: ['cm','cost margin','cost','cm threshold','cm_value'] // optional column
};

// helper: find column header name in sheet headers (case-insensitive, fuzzy by includes)
function findHeader(headerList, headers) {
  const lowerHeaders = headers.map(h => (h||'').toString().toLowerCase().trim());
  for (const candidate of headerList) {
    const lc = candidate.toLowerCase();
    // exact or includes match
    const idx = lowerHeaders.findIndex(h => h === lc || h.includes(lc) || lc.includes(h));
    if (idx !== -1) return { idx, header: headers[idx] };
  }
  return null;
}

function detectColumns(headers) {
  const result = {};
  for (const key of Object.keys(headerCandidates)) {
    const found = findHeader(headerCandidates[key], headers);
    if (found) result[key] = found.idx;
  }
  return result;
}

// configurable matching rules — adjust to your business logic
const config = {
  // if sheet doesn't have a CM column, use this default threshold (numeric)
  defaultCMThreshold: 1000,

  // example rules for size -> container type matching
  // if size contains key, allowed container types array
  typeRules: [
    { sizeContains: '20', allowed: ['20GP','20\'','20ft','20'] },
    { sizeContains: '40', allowed: ['40GP','40\'','40ft','40'] },
    { sizeContains: '40hc', allowed: ['40HC','40\'HC','40ft HC'] }
  ]
};

app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const jobId = uuidv4();
    jobs[jobId] = {
      progress: 0,
      status: 'uploaded',
      result: null,
      error: null
    };

    // Start parse in a "background" async (we still send jobId immediately)
    parseFile(req.file.path, jobId).catch(err => {
      console.error('Parse error', err);
      jobs[jobId].error = err.message || String(err);
      jobs[jobId].status = 'error';
      jobs[jobId].progress = 100;
    });

    res.json({ jobId });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

// SSE endpoint for progress updates
app.get('/progress/:jobId', (req, res) => {
  const { jobId } = req.params;
  const job = jobs[jobId];
  if (!job) return res.status(404).send('Job not found');

  // set headers for SSE
  res.set({
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    Connection: 'keep-alive'
  });
  res.flushHeaders();

  // send current state immediately
  res.write(`event: progress\n`);
  res.write(`data: ${JSON.stringify({ progress: job.progress, status: job.status, error: job.error })}\n\n`);

  const interval = setInterval(() => {
    const j = jobs[jobId];
    if (!j) {
      res.write(`event: done\n`);
      res.write(`data: ${JSON.stringify({ progress: 100, status: 'missing' })}\n\n`);
      clearInterval(interval);
      res.end();
      return;
    }
    res.write(`event: progress\n`);
    res.write(`data: ${JSON.stringify({ progress: j.progress, status: j.status, error: j.error })}\n\n`);
    if (j.progress >= 100) {
      clearInterval(interval);
      res.end();
    }
  }, 700);

  // close
  req.on('close', () => clearInterval(interval));
});

// after parsing finished, client can fetch results
app.get('/result/:jobId', (req, res) => {
  const job = jobs[req.params.jobId];
  if (!job) return res.status(404).json({ error: 'Job not found' });
  if (job.error) return res.status(500).json({ error: job.error });
  if (job.status !== 'done') return res.status(202).json({ status: job.status });
  res.json({ result: job.result });
});

// parse routine
async function parseFile(filePath, jobId) {
  jobs[jobId].status = 'parsing';
  jobs[jobId].progress = 5;

  // load workbook
  const wb = XLSX.readFile(filePath, { cellDates: true });
  jobs[jobId].progress = 12;
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];

  // convert to JSON rows (array of objects) preserving headers
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (raw.length < 1) {
    jobs[jobId].error = 'Empty sheet';
    jobs[jobId].status = 'error';
    jobs[jobId].progress = 100;
    fs.unlinkSync(filePath);
    return;
  }

  const headers = raw[0].map(h => (h == null ? '' : h.toString()));
  const cols = detectColumns(headers);
  jobs[jobId].progress = 20;

  // We will iterate rows and process
  const rows = [];
  const totalRows = raw.length - 1;
  for (let r = 1; r < raw.length; r++) {
    const rowArr = raw[r];
    const rowObj = {};

    // helper to safely get cell by detected column index
    function cellByKey(key) {
      if (cols[key] == null) return '';
      return rowArr[cols[key]] !== undefined ? rowArr[cols[key]] : '';
    }

    const pol = cellByKey('POL') || '';
    const pod = cellByKey('POD') || '';
    const ratesCell = cellByKey('RATES') || cellByKey('RATE') || '';
    const size = (cellByKey('SIZE') || '').toString();
    const ctype = (cellByKey('CONTAINER_TYPE') || '').toString();

    // parse numeric rate
    const rate = parseFloat(String(ratesCell).toString().replace(/[^0-9.\-]+/g, '')) || 0;

    // determine CM threshold: prefer a CM column in sheet if present
    let cmValue = null;
    if (cols['CM'] != null) {
      const cmCell = rowArr[cols['CM']];
      cmValue = parseFloat(String(cmCell || '').replace(/[^0-9.\-]+/g, '')) || null;
    }

    // if no CM in sheet, use default threshold from config
    const threshold = cmValue != null ? cmValue : config.defaultCMThreshold;

    const belowCM = rate < threshold;

    // type matching: simple rule: check size string contains rule key and container type is in allowed list
    let typeMatch = false;
    for (const rule of config.typeRules) {
      if (size.toString().toLowerCase().includes(rule.sizeContains.toLowerCase())) {
        // normalize container type string
        const cNormalized = (ctype || '').toString().toUpperCase();
        if (rule.allowed.some(a => cNormalized.includes(a.toUpperCase()))) {
          typeMatch = true;
          break;
        }
      }
    }

    rowObj._rowNumber = r + 1;
    rowObj.POL = pol;
    rowObj.POD = pod;
    rowObj.RATE = rate;
    rowObj.SIZE = size;
    rowObj.CONTAINER_TYPE = ctype;
    rowObj.CM_THRESHOLD_USED = threshold;
    rowObj.BELOW_CM = belowCM;
    rowObj.TYPE_MATCH = typeMatch;

    rows.push(rowObj);

    // update progress periodically
    if (r % Math.max(1, Math.floor(totalRows / 20)) === 0) {
      jobs[jobId].progress = Math.min(95, 20 + Math.floor((r / totalRows) * 75));
    }
  }

  // finalize
  jobs[jobId].result = { headersDetected: headers, columnsDetected: cols, rowsCount: rows.length, rows };
  jobs[jobId].status = 'done';
  jobs[jobId].progress = 100;

  // cleanup file
  try { fs.unlinkSync(filePath); } catch (e) {}
}

// start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server listening on ${PORT}`));
