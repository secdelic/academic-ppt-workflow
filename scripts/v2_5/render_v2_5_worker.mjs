import readline from "node:readline";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

// One short-lived worker is owned by one batch-regression invocation. It has
// no socket, server, external access, or persistent state beyond the process.
// Dynamic module instances keep each deck isolated while Node's package cache
// avoids reloading PptxGenJS for every deck.
const rendererUrl = pathToFileURL(path.resolve(path.dirname(fileURLToPath(import.meta.url)), "render_v2_5_deck.mjs"));
const input = readline.createInterface({ input: process.stdin, crlfDelay: Infinity });
let jobIndex = 0;

for await (const line of input) {
  if (!line.trim()) continue;
  let job = null;
  try {
    job = JSON.parse(line);
    if (!job.job_id || !job.graph_path || !job.output_path) throw new Error("Worker job contract is incomplete");
    process.argv = [process.argv[0], "render_v2_5_deck.mjs", job.graph_path, job.output_path];
    const cacheKey = `${jobIndex++}-${encodeURIComponent(job.job_id)}`;
    await import(`${rendererUrl.href}?batch_job=${cacheKey}`);
    console.log(`V25_WORKER_RESULT=${JSON.stringify({ job_id: job.job_id, ok: true })}`);
  } catch (error) {
    console.log(`V25_WORKER_RESULT=${JSON.stringify({ job_id: job?.job_id || null, ok: false, error: String(error?.stack || error) })}`);
  }
}
