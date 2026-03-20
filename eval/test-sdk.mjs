/**
 * Quick test of SDK builder pipeline
 */
import { CopilotClient, approveAll } from "@github/copilot-sdk";
import { join } from "path";
import { CLI_PATH, EVAL_OUTPUT_ROOT, SKILLS_DIR } from "./lib/runtime/environment.mjs";

const OUTPUT_DIR = EVAL_OUTPUT_ROOT;

async function main() {
  const c = new CopilotClient();
  await c.start();
  console.log("Client connected");

  const session = await c.createSession({
    model: "claude-sonnet-4.5",
    onPermissionRequest: approveAll,
  });
  console.log("Session created:", session.id);

  const prompt = `You are a slide builder. Read the design skill files, then build one slide.

FIRST: Read these files for design guidance:
- ${join(SKILLS_DIR, "slide-design-principles.md")}

CLI TOOL: ${CLI_PATH}
RULES: Use --color not --font-color. Use --alignment not --horizontal-alignment. Don't use \\n in --text. pptcli service is already running. Close existing sessions first via session list + session close.

TASK: Build a 3-card KPI dashboard slide.
Content: Revenue $45M up 12%, Costs $32M on budget, Margin 29% up 3pp.
Use Corporate Blue palette (Primary #0B3D91, Accent #FF6B35, Positive #2E8B57).
Use the KPI Card Dashboard archetype with 3-across layout:
  Card 1: x=36, y=100, w=276, h=160
  Card 2: x=332, y=100, w=276, h=160
  Card 3: x=628, y=100, w=276, h=160

PPTX: ${join(OUTPUT_DIR, "sdk-test.pptx")}
PNG: ${join(OUTPUT_DIR, "sdk-test.png")}

Build it now, export PNG, close and save.`;

  console.log("Sending prompt (%d chars)...", prompt.length);
  const start = Date.now();
  const resp = await session.sendAndWait({ prompt }, 300000); // 5 min timeout
  const elapsed = ((Date.now() - start) / 1000).toFixed(1);
  console.log(`Response in ${elapsed}s (${resp?.data?.content?.length} chars)`);
  console.log(resp?.data?.content?.slice(0, 500));

  await session.destroy();
  await c.stop();
  console.log("Done");
}

main().catch((e) => {
  console.error("Error:", e.message);
  process.exit(1);
});
