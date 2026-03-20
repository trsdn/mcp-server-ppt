export const FAILURE_CATEGORIES = Object.freeze({
  schemaError: "schema_error",
  timeout: "timeout",
  artifactMissing: "artifact_missing",
  artifactInvalid: "artifact_invalid",
  toolFailure: "tool_failure",
  reviewInvalid: "review_invalid",
  cleanupFailure: "cleanup_failure",
  runtimeUnavailable: "runtime_unavailable",
  invalidConfiguration: "invalid_configuration",
  transportMismatch: "transport_mismatch",
  interrupted: "interrupted",
});

export const FAILURE_DISPOSITIONS = Object.freeze({
  retryable: "retryable",
  fatal: "fatal",
});

export const CONTRACTS = Object.freeze({
  evaluationRequest: "evaluation-request/v1",
  builderSummaryResponse: "builder-summary/v1",
  builderCarryoverEntry: "builder-carryover-entry/v1",
  judgmentRequest: "judgment-request/v1",
  judgeResponse: "judge-response/v1",
  reviewerCarryoverEntry: "reviewer-carryover-entry/v1",
});

export const JUDGE_DIMENSION_KEYS = Object.freeze([
  "archetypeMatch",
  "variantMatch",
  "actionTitle",
  "infoHierarchy",
  "contentDensity",
  "zoneCompliance",
  "sourceCitations",
  "evidenceSupport",
  "visualExecution",
]);

function compactText(value, maxLength = 240) {
  return String(value || "").replace(/\s+/g, " ").trim().slice(0, maxLength);
}

function normalizeStringArray(value, maxItems = 8, maxLength = 180) {
  return Array.isArray(value)
    ? value
      .map((item) => compactText(item, maxLength))
      .filter(Boolean)
      .slice(0, maxItems)
    : [];
}

function normalizeDimensionScores(dimensionScores) {
  return Object.fromEntries(
    JUDGE_DIMENSION_KEYS.map((key) => {
      const dimension = dimensionScores?.[key] || {};
      return [
        key,
        {
          score: Number.isInteger(dimension.score) ? dimension.score : null,
          reason: compactText(dimension.reason || "", 240) || null,
        },
      ];
    })
  );
}

function normalizeValidation(validation) {
  if (!validation || typeof validation !== "object") {
    return null;
  }

  return {
    status: compactText(validation.status || "", 80) || null,
    contract: compactText(validation.contract || "", 120) || null,
    source: compactText(validation.source || "", 80) || null,
    failureCategory: compactText(validation.failure?.category || "", 120) || null,
  };
}

export function createProtocolFailure(category, message, details = {}) {
  return {
    category,
    message,
    ...details,
  };
}

export function classifyExecutionFailure(error, { fallbackCategory = FAILURE_CATEGORIES.toolFailure } = {}) {
  const rawMessage = error instanceof Error ? error.message : String(error || "");
  const message = rawMessage.toLowerCase();
  let category = fallbackCategory;

  if (message.includes("timeout")) {
    category = FAILURE_CATEGORIES.timeout;
  } else if (message.includes("artifact") && message.includes("missing")) {
    category = FAILURE_CATEGORIES.artifactMissing;
  } else if (
    message.includes("artifact") && (
      message.includes("invalid")
      || message.includes("signature")
      || message.includes("too small")
      || message.includes("not stable")
      || message.includes("not a file")
    )
  ) {
    category = FAILURE_CATEGORIES.artifactInvalid;
  } else if (
    message.includes("unsupported transport")
    || (message.includes("transport") && message.includes("instruction"))
    || message.includes("incompatible with instructions")
  ) {
    category = FAILURE_CATEGORIES.transportMismatch;
  } else if (
    message.includes("invalid configuration")
    || message.includes("not implemented yet")
    || message.includes("missing instructions")
    || message.includes("unsupported eval runtime")
  ) {
    category = FAILURE_CATEGORIES.invalidConfiguration;
  } else if (
    message.includes("invalid worker output")
    || message.includes("session closed")
    || message.includes("transport closed")
    || message.includes("connection reset")
    || message.includes("socket hang up")
    || message.includes("service start")
    || message.includes("service unavailable")
    || message.includes("econnreset")
    || message.includes("epipe")
  ) {
    category = FAILURE_CATEGORIES.runtimeUnavailable;
  } else if (message.includes("cleanup")) {
    category = FAILURE_CATEGORIES.cleanupFailure;
  } else if (message.includes("cancelled") || message.includes("canceled") || message.includes("interrupted")) {
    category = FAILURE_CATEGORIES.interrupted;
  }

  const retryable = (
    category === FAILURE_CATEGORIES.timeout
    || category === FAILURE_CATEGORIES.runtimeUnavailable
    || category === FAILURE_CATEGORIES.cleanupFailure
    || category === FAILURE_CATEGORIES.artifactMissing
    || category === FAILURE_CATEGORIES.artifactInvalid
    || category === FAILURE_CATEGORIES.schemaError
    || category === FAILURE_CATEGORIES.reviewInvalid
    || (category === FAILURE_CATEGORIES.toolFailure && (
      message.includes("busy")
      || message.includes("temporar")
      || message.includes("try again")
      || message.includes("retry")
      || message.includes("throttle")
    ))
  );

  return {
    category,
    disposition: retryable ? FAILURE_DISPOSITIONS.retryable : FAILURE_DISPOSITIONS.fatal,
    retryable,
    message: rawMessage || "Unknown failure",
  };
}

export function createFailureDetails(error, options = {}) {
  return classifyExecutionFailure(error, options);
}

export function categorizeExecutionFailure(error, fallbackCategory = FAILURE_CATEGORIES.toolFailure) {
  return classifyExecutionFailure(error, { fallbackCategory }).category;
}

export function createEvaluationRequestEnvelope({
  promptId = "",
  prompt = "",
  archetype = "",
  archetypeFamily = null,
  expectedVariant = null,
  transport = "cli",
  pngPath = "",
  pptxPath = "",
  builderCarryover = [],
  reviewerCarryover = [],
} = {}) {
  return {
    contract: CONTRACTS.evaluationRequest,
    payload: {
      promptId,
      prompt,
      archetype,
      archetypeFamily: archetypeFamily || archetype,
      expectedVariant: expectedVariant || null,
      transport,
      pngPath,
      pptxPath,
      builderCarryover,
      reviewerCarryover,
    },
  };
}

export function createJudgmentRequestEnvelope({
  promptId = "",
  prompt = "",
  archetype = "",
  archetypeFamily = null,
  expectedVariant = null,
  pngPath = "",
  builderCarryover = [],
  reviewerCarryover = [],
} = {}) {
  return {
    contract: CONTRACTS.judgmentRequest,
    payload: {
      promptId,
      prompt,
      archetype,
      archetypeFamily: archetypeFamily || archetype,
      expectedVariant: expectedVariant || null,
      pngPath,
      builderCarryover,
      reviewerCarryover,
    },
  };
}

export function createBuilderCarryoverEntry({
  loopNumber = null,
  promptId = "",
  prompt = "",
  summary = null,
  validation = null,
} = {}) {
  return {
    contract: CONTRACTS.builderCarryoverEntry,
    payload: {
      loopNumber: Number.isInteger(loopNumber) ? loopNumber : null,
      promptId: compactText(promptId, 160) || null,
      prompt: compactText(prompt, 240) || null,
      summary: summary
        ? {
          archetype: compactText(summary.archetype || "unknown", 120) || "unknown",
          palette: compactText(summary.palette || "unknown", 120) || "unknown",
          shapeCount: Number.isInteger(summary.shapeCount) ? summary.shapeCount : null,
          preservedFacts: normalizeStringArray(summary.preservedFacts, 8, 180),
          rationale: compactText(summary.rationale || "", 240) || null,
        }
        : null,
      validation: normalizeValidation(validation),
    },
  };
}

export function createReviewerCarryoverEntry({
  loopNumber = null,
  promptId = "",
  prompt = "",
  judgment = null,
  validation = null,
} = {}) {
  return {
    contract: CONTRACTS.reviewerCarryoverEntry,
    payload: {
      loopNumber: Number.isInteger(loopNumber) ? loopNumber : null,
      promptId: compactText(promptId, 160) || null,
      prompt: compactText(prompt, 240) || null,
      judgment: judgment
        ? {
          archetypeUsed: compactText(judgment.archetypeUsed || "unknown", 120) || "unknown",
          archetypeExpected: compactText(judgment.archetypeExpected || "unknown", 120) || "unknown",
          summary: compactText(judgment.summary || "", 240) || null,
          dimensionScores: normalizeDimensionScores(judgment.dimensionScores),
          totalScore: Number.isInteger(judgment.totalScore) ? judgment.totalScore : null,
          maxScore: Number.isInteger(judgment.maxScore) ? judgment.maxScore : null,
          gaps: normalizeStringArray(judgment.gaps, 8, 180),
        }
        : null,
      validation: normalizeValidation(validation),
    },
  };
}

export function getBuilderSummaryResponseSchemaExample() {
  return {
    contract: CONTRACTS.builderSummaryResponse,
    payload: {
      archetype: "string",
      palette: "string",
      shapeCount: 0,
      preservedFacts: ["string"],
      rationale: "string",
    },
  };
}

export function getJudgeResponseSchemaExample() {
  return {
    contract: CONTRACTS.judgeResponse,
    payload: {
      prompt: "string",
      archetypeUsed: "string",
      archetypeExpected: "string",
      summary: "string",
      dimensionScores: Object.fromEntries(
        JUDGE_DIMENSION_KEYS.map((key) => [key, { score: 0, reason: "string" }])
      ),
      totalScore: 0,
      maxScore: JUDGE_DIMENSION_KEYS.length * 2,
      gaps: ["string"],
    },
  };
}

export function formatProtocolExample(value) {
  return JSON.stringify(value, null, 2);
}
