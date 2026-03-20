import {
  createCopilotSdkRuntime,
  destroyCopilotSdkRuntime,
  executeCopilotSdkFreshRequest,
  executeCopilotSdkSessionRequest,
} from "./copilot-sdk-runtime.mjs";
import { runIsolatedWorkerRequest } from "./isolated-worker-runtime.mjs";
import {
  executeWithRetry,
  loadInstructionsFile,
  normalizeTransport,
  validateTransportInstructionsCompatibility,
  verifyArtifactFile,
  verifyBuildArtifacts,
  waitForArtifactFile,
} from "./reliability.mjs";

function resolveRuntimeKind(agent, options = {}) {
  return options.runtimeKind || agent?.runtime || agent?.runtimeKind || "copilot-sdk";
}

function resolveExecutionMode(agent, options = {}) {
  if (options.executionMode) return options.executionMode;
  if (agent?.isolatedProcess === true || agent?.executionMode === "isolated-process") return "isolated-process";
  if (agent?.reuseSessionContext === true || agent?.executionMode === "reuse-session") return "reuse-session";
  return "fresh-session";
}

function getRuntimeImplementation(runtimeKind) {
  switch (runtimeKind) {
    case "copilot-sdk":
      return {
        createSessionRuntime: createCopilotSdkRuntime,
        destroySessionRuntime: destroyCopilotSdkRuntime,
        executeFreshRequest: executeCopilotSdkFreshRequest,
        executeSessionRequest: executeCopilotSdkSessionRequest,
      };
    case "acp":
      throw new Error("ACP runtime is not implemented yet. Configure runtime='copilot-sdk' or omit the runtime field.");
    default:
      throw new Error(`Unsupported eval runtime '${runtimeKind}'.`);
  }
}

export function shouldUseIsolatedProcess(agent) {
  return resolveExecutionMode(agent) === "isolated-process";
}

export function shouldReuseSessionContext(agent) {
  return resolveExecutionMode(agent) === "reuse-session";
}

export async function createRuntime(agent, options = {}) {
  const runtimeKind = resolveRuntimeKind(agent, options);
  const executionMode = resolveExecutionMode(agent, options);
  const implementation = getRuntimeImplementation(runtimeKind);

  const runtime = {
    agent: { ...agent, runtime: runtimeKind },
    runtimeKind,
    executionMode,
    implementation,
    sessionRuntime: null,
  };

  if (executionMode === "reuse-session") {
    runtime.sessionRuntime = await implementation.createSessionRuntime(runtime.agent);
  }

  return runtime;
}

export async function destroyRuntime(runtime) {
  if (!runtime?.sessionRuntime) return;
  await runtime.implementation.destroySessionRuntime(runtime.sessionRuntime);
}

export async function executeRuntimeRequest(runtime, request) {
  if (runtime.executionMode === "isolated-process") {
    return runIsolatedWorkerRequest(
      {
        ...request,
        agent: runtime.agent,
      },
      request.workerTimeoutMs || request.timeoutMs || 30000
    );
  }

  if (runtime.executionMode === "reuse-session") {
    return runtime.implementation.executeSessionRequest(runtime.sessionRuntime, request);
  }

  return runtime.implementation.executeFreshRequest(runtime.agent, request);
}

export async function executeAgentRequest(agent, request, options = {}) {
  const runtime = await createRuntime(agent, options);
  try {
    return await executeRuntimeRequest(runtime, request);
  } finally {
    await destroyRuntime(runtime);
  }
}

export {
  executeWithRetry,
  loadInstructionsFile,
  normalizeTransport,
  validateTransportInstructionsCompatibility,
  verifyArtifactFile,
  verifyBuildArtifacts,
  waitForArtifactFile,
};

export { ARCHETYPES_DIR, resolveArchetypeFamily } from "./environment.mjs";
