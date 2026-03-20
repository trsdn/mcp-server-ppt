import { closeSync, existsSync, openSync, readFileSync, readSync, statSync } from "fs";
import { basename, join } from "path";
import {
  FAILURE_CATEGORIES,
  createFailureDetails,
} from "../protocol/index.mjs";

const SUPPORTED_TRANSPORTS = new Set(["cli", "mcp"]);
const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47];
const ZIP_SIGNATURES = [
  [0x50, 0x4b, 0x03, 0x04],
  [0x50, 0x4b, 0x05, 0x06],
  [0x50, 0x4b, 0x07, 0x08],
];

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function getDefaultMinBytes(kind) {
  switch (kind) {
    case "pptx":
      return 512;
    case "png":
    default:
      return 128;
  }
}

function readHeaderBytes(filePath, length = 8) {
  let fd;
  try {
    fd = openSync(filePath, "r");
    const buffer = Buffer.alloc(length);
    const bytesRead = readSync(fd, buffer, 0, length, 0);
    return Array.from(buffer.subarray(0, bytesRead));
  } finally {
    if (fd !== undefined) {
      closeSync(fd);
    }
  }
}

function matchesSignature(actual, expected) {
  if (actual.length < expected.length) return false;
  return expected.every((byte, index) => actual[index] === byte);
}

function detectInstructionsTransport(instructionsText = "", instructionsFile = "") {
  const normalizedText = instructionsText.toLowerCase();
  const normalizedFile = basename(instructionsFile || "").toLowerCase();
  const cliSignals = [];
  const mcpSignals = [];

  if (normalizedFile.includes("mcp")) mcpSignals.push("instructions file name");
  if (normalizedFile.includes("cli")) cliSignals.push("instructions file name");

  if (normalizedText.includes("pptcli")) cliSignals.push("pptcli mention");
  if (normalizedText.includes("use the cli")) cliSignals.push("CLI directive");
  if (normalizedText.includes("session create")) cliSignals.push("CLI session workflow");

  if (normalizedText.includes("powerpoint mcp")) mcpSignals.push("MCP mention");
  if (normalizedText.includes("mcp server")) mcpSignals.push("MCP server mention");
  if (normalizedText.includes("file(action:")) mcpSignals.push("MCP tool recipe");

  if (cliSignals.length && !mcpSignals.length) {
    return { inferredTransport: "cli", evidence: cliSignals };
  }

  if (mcpSignals.length && !cliSignals.length) {
    return { inferredTransport: "mcp", evidence: mcpSignals };
  }

  if (mcpSignals.length && cliSignals.length) {
    return {
      inferredTransport: "mixed",
      evidence: [...mcpSignals, ...cliSignals],
    };
  }

  return {
    inferredTransport: "unknown",
    evidence: [],
  };
}

export function normalizeTransport(transport = "cli") {
  const normalized = String(transport || "cli").trim().toLowerCase();
  if (!SUPPORTED_TRANSPORTS.has(normalized)) {
    throw new Error(`Invalid configuration: unsupported transport '${transport}'. Expected one of: cli, mcp.`);
  }

  return normalized;
}

export function validateTransportInstructionsCompatibility({
  transport = "cli",
  instructionsFile = "",
  instructionsText = "",
} = {}) {
  const normalizedTransport = normalizeTransport(transport);
  const detected = detectInstructionsTransport(instructionsText, instructionsFile);

  if (detected.inferredTransport === "mixed") {
    return {
      ok: false,
      category: FAILURE_CATEGORIES.transportMismatch,
      message: `Transport '${normalizedTransport}' is incompatible with mixed instructions in '${instructionsFile}'.`,
      transport: normalizedTransport,
      ...detected,
    };
  }

  if (
    detected.inferredTransport !== "unknown"
    && detected.inferredTransport !== normalizedTransport
  ) {
    return {
      ok: false,
      category: FAILURE_CATEGORIES.transportMismatch,
      message: `Transport '${normalizedTransport}' is incompatible with instructions '${instructionsFile}' inferred for '${detected.inferredTransport}'.`,
      transport: normalizedTransport,
      ...detected,
    };
  }

  return {
    ok: true,
    transport: normalizedTransport,
    ...detected,
  };
}

export function loadInstructionsFile({
  baseDir,
  instructionsFile,
  expectedTransport = null,
  label = "instructions",
} = {}) {
  if (!instructionsFile) {
    return {
      ok: false,
      category: FAILURE_CATEGORIES.invalidConfiguration,
      message: `Missing ${label} file configuration.`,
    };
  }

  const fullPath = join(baseDir, instructionsFile);
  if (!existsSync(fullPath)) {
    return {
      ok: false,
      category: FAILURE_CATEGORIES.invalidConfiguration,
      message: `Missing ${label} file '${instructionsFile}'.`,
      path: fullPath,
    };
  }

  const text = readFileSync(fullPath, "utf-8");
  const compatibility = expectedTransport
    ? validateTransportInstructionsCompatibility({
      transport: expectedTransport,
      instructionsFile,
      instructionsText: text,
    })
    : { ok: true };

  if (!compatibility.ok) {
    return {
      ok: false,
      category: compatibility.category,
      message: compatibility.message,
      path: fullPath,
      compatibility,
    };
  }

  return {
    ok: true,
    path: fullPath,
    text,
    compatibility,
  };
}

export function verifyArtifactFile(filePath, { kind = "png", minBytes } = {}) {
  if (!existsSync(filePath)) {
    return {
      ok: false,
      category: FAILURE_CATEGORIES.artifactMissing,
      message: `${kind.toUpperCase()} artifact missing at '${filePath}'.`,
      path: filePath,
    };
  }

  try {
    const stats = statSync(filePath);
    if (!stats.isFile()) {
      return {
        ok: false,
        category: FAILURE_CATEGORIES.artifactInvalid,
        message: `${kind.toUpperCase()} artifact at '${filePath}' is not a file.`,
        path: filePath,
      };
    }

    const minimumSize = minBytes || getDefaultMinBytes(kind);
    if (stats.size < minimumSize) {
      return {
        ok: false,
        category: FAILURE_CATEGORIES.artifactInvalid,
        message: `${kind.toUpperCase()} artifact at '${filePath}' is too small (${stats.size} bytes).`,
        path: filePath,
        size: stats.size,
      };
    }

    const header = readHeaderBytes(filePath);
    const signatureOk = kind === "pptx"
      ? ZIP_SIGNATURES.some((signature) => matchesSignature(header, signature))
      : matchesSignature(header, PNG_SIGNATURE);

    if (!signatureOk) {
      return {
        ok: false,
        category: FAILURE_CATEGORIES.artifactInvalid,
        message: `${kind.toUpperCase()} artifact at '${filePath}' has an invalid file signature.`,
        path: filePath,
      };
    }

    return {
      ok: true,
      path: filePath,
      kind,
      size: stats.size,
      mtimeMs: stats.mtimeMs,
    };
  } catch (error) {
    const failure = createFailureDetails(error, {
      fallbackCategory: FAILURE_CATEGORIES.artifactInvalid,
    });
    return {
      ok: false,
      ...failure,
      path: filePath,
    };
  }
}

export async function waitForArtifactFile(filePath, {
  kind = "png",
  timeoutMs = 2500,
  settleMs = 350,
  minBytes,
} = {}) {
  const startedAt = Date.now();
  let lastFailure = {
    ok: false,
    category: FAILURE_CATEGORIES.artifactMissing,
    message: `${kind.toUpperCase()} artifact missing at '${filePath}'.`,
    path: filePath,
  };

  while ((Date.now() - startedAt) <= timeoutMs) {
    const firstPass = verifyArtifactFile(filePath, { kind, minBytes });
    if (firstPass.ok) {
      await sleep(settleMs);
      const secondPass = verifyArtifactFile(filePath, { kind, minBytes });
      if (
        secondPass.ok
        && firstPass.size === secondPass.size
        && firstPass.mtimeMs === secondPass.mtimeMs
      ) {
        return {
          ok: true,
          path: filePath,
          kind,
          size: secondPass.size,
          stable: true,
        };
      }

      lastFailure = {
        ok: false,
        category: FAILURE_CATEGORIES.artifactInvalid,
        message: `${kind.toUpperCase()} artifact at '${filePath}' is not stable yet.`,
        path: filePath,
      };
    } else {
      lastFailure = firstPass;
    }

    await sleep(Math.min(250, Math.max(100, Math.floor(settleMs / 2))));
  }

  return lastFailure;
}

export async function verifyBuildArtifacts({
  pngPath,
  pptxPath,
  requirePptx = true,
  timeoutMs = 2500,
} = {}) {
  const png = await waitForArtifactFile(pngPath, { kind: "png", timeoutMs });
  if (!png.ok) return png;

  if (!requirePptx) {
    return {
      ok: true,
      artifacts: {
        png,
      },
    };
  }

  const pptx = await waitForArtifactFile(pptxPath, { kind: "pptx", timeoutMs });
  if (!pptx.ok) return pptx;

  return {
    ok: true,
    artifacts: {
      png,
      pptx,
    },
  };
}

export async function executeWithRetry(operation, {
  maxAttempts = 1,
  baseDelayMs = 1500,
  maxDelayMs = 6000,
  isSuccess = (result) => Boolean(result?.ok ?? result?.success),
  classifyFailure = (result) => createFailureDetails(result?.error, {
    fallbackCategory: result?.errorCategory,
  }),
  shouldRetry = ({ failure }) => Boolean(failure?.retryable),
  onRetry,
} = {}) {
  const history = [];

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    let result;

    try {
      result = await operation({ attempt, maxAttempts });
    } catch (error) {
      const failure = createFailureDetails(error);
      result = {
        ok: false,
        error: failure.message,
        errorCategory: failure.category,
        failure,
      };
    }

    if (isSuccess(result)) {
      return {
        ...result,
        retry: {
          attempts: attempt,
          recovered: attempt > 1,
          exhausted: false,
          history,
        },
      };
    }

    const failure = result?.failure || classifyFailure(result);
    history.push({
      attempt,
      category: failure?.category || result?.errorCategory || null,
      disposition: failure?.disposition || null,
      retryable: Boolean(failure?.retryable),
      message: failure?.message || result?.error || "Unknown failure",
    });

    const canRetry = attempt < maxAttempts && shouldRetry({
      result,
      failure,
      attempt,
      maxAttempts,
      history,
    });

    if (!canRetry) {
      return {
        ...result,
        failure,
        retry: {
          attempts: attempt,
          recovered: false,
          exhausted: attempt >= maxAttempts && Boolean(failure?.retryable),
          history,
        },
      };
    }

    const delayMs = Math.min(maxDelayMs, baseDelayMs * (2 ** (attempt - 1)));
    await onRetry?.({
      result,
      failure,
      attempt,
      maxAttempts,
      delayMs,
      history,
    });
    await sleep(delayMs);
  }

  return {
    ok: false,
    error: "Retry policy exhausted without producing a result.",
    errorCategory: FAILURE_CATEGORIES.toolFailure,
    failure: createFailureDetails("Retry policy exhausted without producing a result."),
    retry: {
      attempts: maxAttempts,
      recovered: false,
      exhausted: true,
      history,
    },
  };
}
