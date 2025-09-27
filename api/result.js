export default function handler(req, res) {
  const { jobId } = req.query;
  if (!jobId) return res.status(400).json({ error: "Missing jobId" });

  const job = jobs[jobId];
  if (!job) return res.status(404).json({ error: "Job not found" });

  res.status(200).json({ result: job.result });
}
