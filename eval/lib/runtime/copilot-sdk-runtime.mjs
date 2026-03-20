import { CopilotClient, approveAll } from "@github/copilot-sdk";
import { execSync } from "child_process";
import { existsSync } from "fs";
import { CLI_PATH, MCP_SERVER_PATH, REPO_ROOT } from "./environment.mjs";
import { verifyBuildArtifacts, waitForArtifactFile } from "./reliability.mjs";

const BUILD_POLL_INTERVAL_MS = 5000;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function getErrorMessage(error) {
  return error instanceof Error ? error.message : String(error);
}

export function buildCopilotSessionConfig(agent) {
  const sessionConfig = {
    model: agent.model,
    onPermissionRequest: approveAll,
    workingDirectory: REPO_ROOT,
  };

  if (agent.reasoningEffort) {
    sessionConfig.reasoningEffort = agent.reasoningEffort;
  }

  if (agent.transport === "mcp") {
    sessionConfig.mcpServers = {
      ppt: {
        command: MCP_SERVER_PATH,
        args: [],
        tools: ["*"],
        cwd: REPO_ROOT,
      },
    };
    sessionConfig.skillDirectories = [joinSkillsDirectory()];
  }

  return sessionConfig;
}

function joinSkillsDirectory() {
  return `${REPO_ROOT}\\skills`;
}

export async function createCopilotSdkRuntime(agent) {
  const client = new CopilotClient();
  await client.start();
  const session = await client.createSession(buildCopilotSessionConfig(agent));
  return { kind: "copilot-sdk", agent, client, session };
}

export async function destroyCopilotSdkRuntime(runtime, options = {}) {
  if (!runtime) return;

  const { force = false } = options;

  try { await runtime.session?.destroy?.(); } catch {}
  try {
    if (force) {
      await runtime.client?.forceStop?.();
    } else {
      await runtime.client?.stop?.();
    }
  } catch {}
}

export async function executeCopilotSdkFreshRequest(agent, request) {
  if (request.type === "build") {
    return executeFreshBuildRequest(agent, request);
  }

  const runtime = await createCopilotSdkRuntime(agent);
  try {
    return await executeCopilotSdkSessionRequest(runtime, request);
  } catch (error) {
    return { ok: false, error: getErrorMessage(error) };
  } finally {
    await destroyCopilotSdkRuntime(runtime);
  }
}

export async function executeCopilotSdkSessionRequest(runtime, request) {
  switch (request.type) {
    case "prompt":
      return executePromptRequest(runtime, request);
    case "build":
      return executeSessionBuildRequest(runtime, request);
    default:
      throw new Error(`Unsupported Copilot SDK request type: ${request.type}`);
  }
}

async function executePromptRequest(runtime, request) {
  try {
    const response = await runtime.session.sendAndWait({ prompt: request.prompt }, request.timeoutMs);
    return { ok: true, content: response?.data?.content || "" };
  } catch (error) {
    return { ok: false, error: getErrorMessage(error) };
  }
}

async function executeFreshBuildRequest(agent, request) {
  const initialPowerPointPids = getPowerPointProcessIds();
  let runtime = null;
  let result = null;

  try {
    runtime = await createCopilotSdkRuntime(agent);
    result = await executeBuildLoop(runtime, request, {
      beforeFallback: async () => {
        await destroyCopilotSdkRuntime(runtime);
        runtime = null;
      },
      afterSuccess: async () => {
        await destroyCopilotSdkRuntime(runtime);
        runtime = null;
      },
      onError: async () => {
        await destroyCopilotSdkRuntime(runtime, { force: true });
        runtime = null;
      },
    });
    return result;
  } finally {
    if (agent.transport === "mcp") {
      cleanupExtraPowerPointProcesses(initialPowerPointPids);
    }
    if (runtime?.client || runtime?.session) {
      await destroyCopilotSdkRuntime(runtime, { force: true });
    }
  }
}

async function executeSessionBuildRequest(runtime, request) {
  const initialPowerPointPids = getPowerPointProcessIds();

  try {
    return await executeBuildLoop(runtime, request);
  } finally {
    // When reusing session context, the MCP server manages its own PowerPoint
    // lifecycle — don't kill its process between loops.
    if (runtime.agent.transport === "mcp" && !runtime.agent.reuseSessionContext) {
      cleanupExtraPowerPointProcesses(initialPowerPointPids);
    }
  }
}

async function executeBuildLoop(runtime, request, hooks = {}) {
  const session = runtime.session;

  try {
    session.send({ prompt: request.prompt });
    const startedAt = Date.now();

    while ((Date.now() - startedAt) < request.timeoutMs) {
      const artifactStatus = await waitForArtifactFile(request.pngPath, {
        kind: "png",
        timeoutMs: 1500,
      });

      if (artifactStatus.ok) {
        const summaryContent = await requestBuildSummary(session, request);
        if (hooks.afterSuccess) {
          await hooks.afterSuccess();
        }
        return {
          ok: true,
          completion: "png-detected",
          summaryContent,
        };
      }

      await sleep(BUILD_POLL_INTERVAL_MS);
    }

    if (hooks.beforeFallback) {
      await hooks.beforeFallback();
    }

    const delayedArtifact = await verifyBuildArtifacts({
      pngPath: request.pngPath,
      pptxPath: request.pptxPath,
      requirePptx: existsSync(request.pptxPath),
      timeoutMs: 2500,
    });

    if (delayedArtifact.ok) {
      return { ok: true, completion: "png-detected-after-destroy", summaryContent: "" };
    }

    if (existsSync(request.pptxPath) && tryExportFirstSlide(request.pptxPath, request.pngPath)) {
      return { ok: true, completion: "manual-export-after-timeout", summaryContent: "" };
    }

    return { ok: false, error: `Timeout after ${request.timeoutMs}ms waiting for build artifact` };
  } catch (error) {
    if (hooks.onError) {
      await hooks.onError(error);
    }

    const recoveredArtifact = await verifyBuildArtifacts({
      pngPath: request.pngPath,
      pptxPath: request.pptxPath,
      requirePptx: existsSync(request.pptxPath),
      timeoutMs: 2500,
    });

    if (recoveredArtifact.ok) {
      return { ok: true, completion: "png-detected-after-error", summaryContent: "" };
    }

    if (existsSync(request.pptxPath) && tryExportFirstSlide(request.pptxPath, request.pngPath)) {
      return { ok: true, completion: "manual-export-after-error", summaryContent: "" };
    }

    return { ok: false, error: getErrorMessage(error) };
  }
}

async function requestBuildSummary(session, request) {
  if (!request.summaryPrompt) return "";

  const response = await session.sendAndWait(
    { prompt: request.summaryPrompt },
    request.summaryTimeoutMs || 30000
  );

  return response?.data?.content || "";
}

function tryExportFirstSlide(pptxPath, pngPath) {
  try {
    const openOut = execSync(`"${CLI_PATH}" session open "${pptxPath}"`, { encoding: "utf-8", timeout: 15000 });
    const match = openOut.match(/\{[\s\S]*\}/);
    if (!match) return false;
    const data = JSON.parse(match[0]);
    const sessionId = data.sessionId;
    if (!sessionId) return false;

    try {
      execSync(
        `"${CLI_PATH}" export slide-to-image -s ${sessionId} --slide-index 1 --destination-path "${pngPath}" --width 1920 --height 1080`,
        { encoding: "utf-8", timeout: 30000 }
      );
    } finally {
      try {
        execSync(`"${CLI_PATH}" session close -s ${sessionId} --save`, { encoding: "utf-8", timeout: 15000 });
      } catch {}
    }

    return existsSync(pngPath);
  } catch {
    return false;
  }
}

function getPowerPointProcessIds() {
  try {
    const out = execSync(
      "powershell -NoProfile -Command \"Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id\"",
      { encoding: "utf-8", timeout: 10000 }
    );

    return new Set(
      out
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter(Boolean)
        .map((id) => Number.parseInt(id, 10))
        .filter(Number.isFinite)
    );
  } catch {
    return new Set();
  }
}

function cleanupExtraPowerPointProcesses(initialPids) {
  const currentPids = getPowerPointProcessIds();
  for (const pid of currentPids) {
    if (!initialPids.has(pid)) {
      try {
        execSync(`powershell -NoProfile -Command "Stop-Process -Id ${pid}"`, { encoding: "utf-8", timeout: 10000 });
      } catch {}
    }
  }
}
