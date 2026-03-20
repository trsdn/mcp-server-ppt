import { existsSync } from "fs";
import { CopilotClient, approveAll } from "@github/copilot-sdk";
import { REPO_ROOT, SKILLS_ROOT, DEFAULT_MCP_SERVER_PATH, MCP_SERVER_ENV_KEYS } from "./constants.mjs";

function getErrorMessage(error) {
  return error instanceof Error ? error.message : String(error || "Unknown error");
}

export function resolveMcpServerPath(explicitPath = null) {
  if (explicitPath) {
    return explicitPath;
  }

  for (const key of MCP_SERVER_ENV_KEYS) {
    if (process.env[key]) {
      return process.env[key];
    }
  }

  return DEFAULT_MCP_SERVER_PATH;
}

function createSessionConfig({
  model,
  mcpServerPath,
  enableMcp = true,
  enableSkills = true,
}) {
  const config = {
    model,
    onPermissionRequest: approveAll,
    workingDirectory: REPO_ROOT,
  };

  if (enableSkills) {
    config.skillDirectories = [SKILLS_ROOT];
  }

  if (enableMcp) {
    config.mcpServers = {
      ppt: {
        command: mcpServerPath,
        args: [],
        tools: ["*"],
        cwd: REPO_ROOT,
      },
    };
  }

  return config;
}

export async function createRuntime(options) {
  const client = new CopilotClient();
  await client.start();

  return {
    client,
    options: {
      ...options,
      mcpServerPath: resolveMcpServerPath(options.mcpServerPath),
    },
  };
}

export async function destroyRuntime(runtime) {
  if (!runtime) {
    return;
  }

  try {
    await runtime.client?.stop?.();
  } catch {
    // Best-effort shutdown.
  }
}

function bindPhaseEvents(session, phaseName, verbose) {
  session.on((event) => {
    if (event.type === "assistant.message_delta") {
      process.stdout.write(event.data?.deltaContent || "");
      return;
    }

    if (!verbose) {
      return;
    }

    if (event.type === "tool.call") {
      console.log(`\n[${phaseName}] tool.call ${event.data?.name || "unknown"}`);
      return;
    }

    if (event.type === "tool.result") {
      console.log(`\n[${phaseName}] tool.result`);
      return;
    }

    if (event.type === "session.error") {
      console.log(`\n[${phaseName}] session.error ${event.data?.message || ""}`);
    }
  });
}

export async function runPhase(runtime, phase) {
  const session = await runtime.client.createSession(
    createSessionConfig({
      model: runtime.options.model,
      mcpServerPath: runtime.options.mcpServerPath,
      enableMcp: phase.enableMcp,
      enableSkills: phase.enableSkills ?? true,
    })
  );

  bindPhaseEvents(session, phase.name, runtime.options.verbose);

  console.log(`\n=== ${phase.label} ===`);

  try {
    const response = await session.sendAndWait({ prompt: phase.prompt }, phase.timeoutMs);
    const content = response?.data?.content || "";

    if (content && !content.endsWith("\n")) {
      process.stdout.write("\n");
    }

    return { ok: true, content };
  } catch (error) {
    const errorMessage = getErrorMessage(error);
    return {
      ok: false,
      error: errorMessage,
      isTimeout: errorMessage.includes("Timeout"),
      artifactDetected: Boolean(
        phase.successArtifactPath &&
        existsSync(phase.successArtifactPath)
      ),
    };
  } finally {
    try {
      await session.destroy?.();
    } catch {
      // Best-effort session cleanup.
    }
  }
}
