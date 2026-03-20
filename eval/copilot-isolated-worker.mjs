import { executeCopilotSdkFreshRequest } from "./lib/runtime/copilot-sdk-runtime.mjs";

function readStdin() {
  return new Promise((resolve, reject) => {
    let data = "";
    process.stdin.setEncoding("utf8");
    process.stdin.on("data", (chunk) => { data += chunk; });
    process.stdin.on("end", () => resolve(data));
    process.stdin.on("error", reject);
  });
}

async function main() {
  const raw = await readStdin();
  const request = JSON.parse(raw);
  const requestType = request.type || (request.kind === "build" ? "build" : "prompt");
  const result = await executeCopilotSdkFreshRequest(
    {
      ...request.agent,
      isolatedProcess: false,
      executionMode: "fresh-session",
    },
    {
      ...request,
      type: requestType,
    }
  );

  process.stdout.write(JSON.stringify(result));
}

main().catch((error) => {
  process.stdout.write(JSON.stringify({ ok: false, error: error.message }));
  process.exit(1);
});
