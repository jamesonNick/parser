import multer from "multer";
import nextConnect from "next-connect";
import { v4 as uuidv4 } from "uuid";
import xlsx from "xlsx";

// In-memory job store (works only while function is running)
const jobs = {};

const upload = multer({ storage: multer.memoryStorage() });

const apiRoute = nextConnect({
  onError(error, req, res) {
    res.status(501).json({ error: `Error: ${error.message}` });
  },
  onNoMatch(req, res) {
    res.status(405).json({ error: `Method '${req.method}' not allowed` });
  },
});

apiRoute.use(upload.single("file"));

apiRoute.post((req, res) => {
  try {
    const jobId = uuidv4();

    // Parse Excel directly from buffer
    const workbook = xlsx.read(req.file.buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // Example parsing logic
    const rows = json.map((row, idx) => ({
      _rowNumber: idx + 2,
      POL: row.POL || "",
      POD: row.POD || "",
      RATE: row.RATE || "",
      SIZE: row.SIZE || "",
      CONTAINER_TYPE: row.CONTAINER_TYPE || "",
      BELOW_CM: row.RATE < 1000, // sample rule
      TYPE_MATCH: row.SIZE === row.CONTAINER_TYPE,
    }));

    jobs[jobId] = {
      status: "done",
      progress: 100,
      result: {
        headersDetected: Object.keys(json[0] || {}),
        columnsDetected: Object.keys(json[0] || {}),
        rowsCount: rows.length,
        rows,
      },
    };

    res.status(200).json({ jobId });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

export default apiRoute;
