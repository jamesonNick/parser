export default function handler(req, res) {
  const { jobId } = req.query;
  if (!jobId) return res.status(400).json({ error: "Missing jobId" });

  // In real setup, store jobs in Redis/DB
  res.writeHead(200, {
    "Content-Type": "text/event-stream",
    "Cache-Control": "no-cache",
    Connection: "keep-alive",
  });

  const progress = jobs[jobId] || { status: "unknown", progress: 0 };
  res.write(`data: ${JSON.stringify(progress)}\n\n`);
  res.end();
}
