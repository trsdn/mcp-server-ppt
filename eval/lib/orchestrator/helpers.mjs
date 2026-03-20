import { execSync } from "child_process";
import { CLI_PATH } from "../runtime/environment.mjs";
import { FAILURE_CATEGORIES, createFailureDetails } from "../protocol/index.mjs";

export function cleanupPptSessions() {
  try {
    const output = execSync(`"${CLI_PATH}" session list`, {
      encoding: "utf-8",
      timeout: 10000,
    });
    const match = output.match(/\{[\s\S]*\}/);

    if (!match) {
      return { count: 0, closedSessionIds: [] };
    }

    const data = JSON.parse(match[0]);
    const closedSessionIds = [];

    const failedSessionIds = [];

    for (const session of data.sessions || []) {
      try {
        execSync(`"${CLI_PATH}" session close -s ${session.sessionId}`, {
          encoding: "utf-8",
          timeout: 10000,
        });
        closedSessionIds.push(session.sessionId);
      } catch (error) {
        const failure = createFailureDetails(error, {
          fallbackCategory: FAILURE_CATEGORIES.cleanupFailure,
        });
        failedSessionIds.push({
          sessionId: session.sessionId,
          error: failure.message,
          errorCategory: failure.category,
        });
      }
    }

    return {
      ok: failedSessionIds.length === 0,
      count: closedSessionIds.length,
      closedSessionIds,
      failedSessionIds,
    };
  } catch (error) {
    const failure = createFailureDetails(error, {
      fallbackCategory: FAILURE_CATEGORIES.cleanupFailure,
    });
    return {
      ok: false,
      count: 0,
      closedSessionIds: [],
      failedSessionIds: [],
      error: failure.message,
      errorCategory: failure.category,
    };
  }
}

export function ensureService() {
  try {
    execSync(`"${CLI_PATH}" service start`, {
      encoding: "utf-8",
      timeout: 15000,
    });
    return { ok: true };
  } catch (error) {
    const failure = createFailureDetails(error, {
      fallbackCategory: FAILURE_CATEGORIES.runtimeUnavailable,
    });
    return {
      ok: false,
      error: failure.message,
      errorCategory: failure.category,
      failure,
    };
  }
}

export function compactText(value, maxLength = 240) {
  return (value || "").replace(/\s+/g, " ").trim().slice(0, maxLength);
}

export function formatRollingContext(items, heading) {
  if (!items?.length) return `${heading}: none yet`;
  return `${heading}:\n${JSON.stringify(items, null, 2)}`;
}

export function getBuilderTimeoutMs(builder = {}) {
  if (builder.timeoutMs) return builder.timeoutMs;

  const reasoningEffort = (builder.reasoningEffort || "").toLowerCase();
  const model = (builder.model || "").toLowerCase();

  if (reasoningEffort === "xhigh") return 600000;
  if (reasoningEffort === "high" && model.includes("opus")) return 600000;
  if (reasoningEffort === "high") return 480000;
  if (model.includes("opus")) return 420000;
  return 300000;
}
