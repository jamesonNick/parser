// api/index.js (Vercel-ready)

import express from "express";
import multer from "multer";
import XLSX from "xlsx";
import cors from "cors";
import { v4 as uuidv4 } from "uuid";
import fs from "fs";
import serverless from "serverless-http";

const upload = multer({ dest: "/tmp" }); // ✅ use /tmp for Vercel runtime
const app = express();
app.use(cors());
app.use(express.json());

const jobs = {}; // in-memory job tracker

const headerCandidates = {
  POL: ["pol", "port of loading", "port loading"],
  POD: ["pod", "port of discharge", "port discharge"],
  RATES: ["rate", "rates", "freight", "amount", "price"],
  SIZE: ["size", "container size", "container-size"],
  CONTAINER_TYPE: ["container type", "container", "ctype", "type", "container-type"],
  CM: ["cm", "cost margin", "cost", "cm threshold", "cm_value"]
};

// 🔎 Header detection
function findHeader(headerList, headers) {
  const lowerHeaders = headers.map((h) => (h || "").toString().toLowerCase().trim());
  for (const candidate of headerList) {
    const lc = candidate.toLowerCase();
    const idx = lowerHeaders.findIndex((h) => h === lc || h.includes(lc) || lc.includes(h));
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

// ⚙️ Config
const config = {
  defaultCMThreshold: 1000,
  typeRules: [
    { sizeContains: "20", allowed: ["20GP", "20'", "20ft", "20"] },
    { sizeContains: "40", allowed: ["40GP", "40'", "40ft", "40"] },
    { sizeContains: "40hc", allowed: ["40HC", "40'HC", "40ft HC"] },
  ],
};

// 📤 Upload endpoint
app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    const jobId = uuidv4();
    jobs[jobId] = { progress: 0, status: "uploaded", result: null, error: null };

    parseFile(req.file.path, jobId).catch((err) => {
      console.error("Parse error", err);
      jobs[jobId].error = err.message || String(err);
      jobs[jobId].status = "error";
      jobs[jobId].progress = 100;
    });

    res.json({ jobId });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

// 📊 Progress endpoint (SSE)
app.get("/progress/:jobId", (req, res) => {
  const { jobId } = req.params;
  const job = jobs[jobId];
  if (!job) return res.status(404).send("Job not found");

  res.set({
    "Content-Type": "text/event-stream",
    "Cache-Control": "no-cache",
    Connection: "keep-alive",
  });
  res.flushHeaders();

  res.write(`event: progress\n`);
  res.write(`data: ${JSON.stringify({ progress: job.progress, status: job.status, error: job.error })}\n\n`);

  const interval = setInterval(() => {
    const j = jobs[jobId];
    if (!j) {
      res.write(`event: done\n`);
      res.write(`data: ${JSON.stringify({ progress: 100, status: "missing" })}\n\n`);
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

  req.on("close", () => clearInterval(interval));
});

// 📦 Results endpoint
app.get("/result/:jobId", (req, res) => {
  const job = jobs[req.params.jobId];
  if (!job) return res.status(404).json({ error: "Job not found" });
  if (job.error) return res.status(500).json({ error: job.error });
  if (job.status !== "done") return res.status(202).json({ status: job.status });
  res.json({ result: job.result });
});

// 🔄 Parsing routine
async function parseFile(filePath, jobId) {
  jobs[jobId].status = "parsing";
  jobs[jobId].progress = 5;

  const wb = XLSX.readFile(filePath, { cellDates: true });
  jobs[jobId].progress = 12;
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];

  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (raw.length < 1) {
    jobs[jobId].error = "Empty sheet";
    jobs[jobId].status = "error";
    jobs[jobId].progress = 100;
    fs.unlinkSync(filePath);
    return;
  }

  const headers = raw[0].map((h) => (h == null ? "" : h.toString()));
  const cols = detectColumns(headers);
  jobs[jobId].progress = 20;

  const rows = [];
  const totalRows = raw.length - 1;
  for (let r = 1; r < raw.length; r++) {
    const rowArr = raw[r];
    const rowObj = {};

    function cellByKey(key) {
      if (cols[key] == null) return "";
      return rowArr[cols[key]] !== undefined ? rowArr[cols[key]] : "";
    }

    const pol = cellByKey("POL") || "";
    const pod = cellByKey("POD") || "";
    const ratesCell = cellByKey("RATES") || cellByKey("RATE") || "";
    const size = (cellByKey("SIZE") || "").toString();
    const ctype = (cellByKey("CONTAINER_TYPE") || "").toString();

    const rate = parseFloat(String(ratesCell).replace(/[^0-9.\-]+/g, "")) || 0;

    let cmValue = null;
    if (cols["CM"] != null) {
      const cmCell = rowArr[cols["CM"]];
      cmValue = parseFloat(String(cmCell || "").replace(/[^0-9.\-]+/g, "")) || null;
    }

    const threshold = cmValue != null ? cmValue : config.defaultCMThreshold;
    const belowCM = rate < threshold;

    let typeMatch = false;
    for (const rule of config.typeRules) {
      if (size.toLowerCase().includes(rule.sizeContains.toLowerCase())) {
        const cNormalized = (ctype || "").toUpperCase();
        if (rule.allowed.some((a) => cNormalized.includes(a.toUpperCase()))) {
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

    if (r % Math.max(1, Math.floor(totalRows / 20)) === 0) {
      jobs[jobId].progress = Math.min(95, 20 + Math.floor((r / totalRows) * 75));
    }
  }

  jobs[jobId].result = { headersDetected: headers, columnsDetected: cols, rowsCount: rows.length, rows };
  jobs[jobId].status = "done";
  jobs[jobId].progress = 100;

  try {
    fs.unlinkSync(filePath);
  } catch {}
}

// ✅ Export handler for Vercel
export default serverless(app);
