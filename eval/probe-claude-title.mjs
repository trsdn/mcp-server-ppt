import { CopilotClient, approveAll } from "@github/copilot-sdk";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { existsSync } from "fs";
const __dirname = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = join(__dirname, '..');
const CLI_PATH = join(REPO_ROOT, 'src', 'PptMcp.CLI', 'bin', 'Release', 'net9.0-windows', 'pptcli.exe');
const pptx = join(__dirname, 'output', 'probe-claude-title.pptx');
const png = join(__dirname, 'output', 'probe-claude-title.png');
const prompt = `Use the CLI at ${CLI_PATH}. Build one title slide in ${pptx} and export ${png}. Use this exact workflow: session create -> slide create blank -> add 2 textboxes -> format text -> export slide-to-image -> session close --save. Keep it under 10 commands. Title: Revenue +8% validates Q4 plan. Subtitle: Board review | FY2025 Q4 | $128M revenue. When done, reply DONE only.`;
const c = new CopilotClient();
await c.start();
const s = await c.createSession({ model: 'claude-sonnet-4.5', onPermissionRequest: approveAll });
const t0 = Date.now();
try {
  const resp = await s.sendAndWait({ prompt }, 180000);
  console.log('DONE in', ((Date.now()-t0)/1000).toFixed(1), 's');
  console.log(String(resp?.data?.content || '').slice(0, 400));
} catch (e) {
  console.log('ERROR', e.message);
}
await s.destroy();
await c.stop();
console.log('pptx_exists', existsSync(pptx));
console.log('png_exists', existsSync(png));
