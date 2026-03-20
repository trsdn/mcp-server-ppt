import { spawn } from "child_process";
import { EVAL_ROOT, ISOLATED_WORKER_PATH } from "./environment.mjs";

export function runIsolatedWorkerRequest(request, timeoutMs) {
  return new Promise((resolve) => {
    const child = spawn(process.execPath, [ISOLATED_WORKER_PATH], {
      cwd: EVAL_ROOT,
      stdio: ["pipe", "pipe", "pipe"],
    });

    let stdout = "";
    let stderr = "";
    let settled = false;
    const watchdog = setTimeout(() => {
      if (settled) return;
      settled = true;
      child.kill();
      resolve({ ok: false, error: `Isolated worker timeout after ${timeoutMs}ms` });
    }, timeoutMs);

    child.stdout.on("data", (chunk) => { stdout += chunk.toString(); });
    child.stderr.on("data", (chunk) => { stderr += chunk.toString(); });
    child.on("error", (error) => {
      if (settled) return;
      settled = true;
      clearTimeout(watchdog);
      resolve({ ok: false, error: error.message });
    });
    child.on("close", (code) => {
      if (settled) return;
      settled = true;
      clearTimeout(watchdog);
      try {
        const parsed = JSON.parse(stdout.trim() || "{}");
        if (!parsed.ok && stderr.trim()) {
          parsed.error = parsed.error ? `${parsed.error}\n${stderr.trim()}` : stderr.trim();
        }
        resolve(parsed);
      } catch {
        resolve({ ok: false, error: `Invalid worker output (exit ${code}): ${stderr || stdout}`.trim() });
      }
    });

    child.stdin.end(JSON.stringify(request));
  });
}
