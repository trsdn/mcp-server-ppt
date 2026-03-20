/**
 * Debug SDK session events to see where it gets stuck
 */
import { CopilotClient, approveAll } from "@github/copilot-sdk";

async function main() {
  const c = new CopilotClient();
  await c.start();
  console.log("Connected");

  const session = await c.createSession({
    model: "claude-sonnet-4.5",
    onPermissionRequest: approveAll,
  });
  console.log("Session created");

  // Listen to ALL events
  session.on((event) => {
    if (event.type === "assistant.message_delta") {
      process.stdout.write(event.data?.deltaContent || "");
    } else if (event.type === "assistant.message") {
      console.log("\n[MSG COMPLETE]");
    } else if (event.type === "session.idle") {
      console.log("\n[SESSION IDLE]");
    } else if (event.type === "session.error") {
      console.log("\n[SESSION ERROR]", event.data?.message);
    } else if (event.type === "tool.call") {
      console.log("\n[TOOL CALL]", event.data?.name, event.data?.arguments?.slice?.(0, 100));
    } else if (event.type === "tool.result") {
      console.log("[TOOL RESULT]", String(event.data?.result || "").slice(0, 100));
    } else {
      console.log(`\n[${event.type}]`, JSON.stringify(event.data || {}).slice(0, 150));
    }
  });

  // Simple task that requires file reading
  const prompt = `Read the file C:\\Users\\torstenmahr\\github\\mcp-server-ppt\\skills\\shared\\slide-design-principles.md and tell me how many sections (## headings) are defined in it. Just the count.`;
  
  console.log("Sending...");
  session.send({ prompt });

  // Wait for idle or timeout
  await new Promise((resolve) => setTimeout(resolve, 120000));
  
  await session.destroy();
  await c.stop();
}

main().catch((e) => {
  console.error("Fatal:", e.message);
  process.exit(1);
});
