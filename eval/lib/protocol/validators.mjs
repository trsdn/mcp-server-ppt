import {
  CONTRACTS,
  FAILURE_CATEGORIES,
  JUDGE_DIMENSION_KEYS,
  createProtocolFailure,
} from "./contracts.mjs";

function isPlainObject(value) {
  return value != null && typeof value === "object" && !Array.isArray(value);
}

function compactText(value, maxLength = 240) {
  return String(value || "").replace(/\s+/g, " ").trim().slice(0, maxLength);
}

function normalizeStringArray(value, maxItems = 5, maxLength = 180) {
  return Array.isArray(value)
    ? value
      .map((item) => compactText(item, maxLength))
      .filter(Boolean)
      .slice(0, maxItems)
    : [];
}

function extractJsonObject(text) {
  if (!text || !text.trim()) {
    return { ok: false, error: "Response was empty." };
  }

  const trimmed = text.trim();
  const candidates = [{ source: "raw", text: trimmed }];
  const fenced = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenced?.[1]?.trim()) {
    candidates.push({ source: "fenced", text: fenced[1].trim() });
  }

  let lastError = "No JSON object found in response.";

  for (const candidate of candidates) {
    const firstBrace = candidate.text.indexOf("{");
    const lastBrace = candidate.text.lastIndexOf("}");
    if (firstBrace < 0 || lastBrace <= firstBrace) {
      lastError = "No JSON object found in response.";
      continue;
    }

    const jsonText = candidate.text.slice(firstBrace, lastBrace + 1);

    try {
      const value = JSON.parse(jsonText);
      if (!isPlainObject(value)) {
        lastError = "Parsed JSON value was not an object.";
        continue;
      }

      return {
        ok: true,
        source: candidate.source,
        jsonText,
        value,
      };
    } catch (error) {
      lastError = error instanceof Error ? error.message : String(error);
    }
  }

  return { ok: false, error: lastError };
}

function validateString(value, field, issues, { allowEmpty = false } = {}) {
  if (typeof value !== "string") {
    issues.push({ code: `${field}-type`, message: `${field} must be a string.` });
    return;
  }

  if (!allowEmpty && !value.trim()) {
    issues.push({ code: `${field}-empty`, message: `${field} must not be empty.` });
  }
}

function validateInteger(value, field, issues, { min = Number.MIN_SAFE_INTEGER, max = Number.MAX_SAFE_INTEGER } = {}) {
  if (!Number.isInteger(value)) {
    issues.push({ code: `${field}-type`, message: `${field} must be an integer.` });
    return;
  }

  if (value < min || value > max) {
    issues.push({ code: `${field}-range`, message: `${field} must be between ${min} and ${max}.` });
  }
}

function validateBuilderPayload(payload) {
  const issues = [];

  if (!isPlainObject(payload)) {
    return [{ code: "builder-payload-type", message: "builder payload must be an object." }];
  }

  validateString(payload.archetype, "payload.archetype", issues);
  validateString(payload.palette, "payload.palette", issues);

  if (!(payload.shapeCount === null || Number.isInteger(payload.shapeCount))) {
    issues.push({ code: "payload.shapeCount-type", message: "payload.shapeCount must be an integer or null." });
  }

  if (payload.shapeCount != null && payload.shapeCount < 0) {
    issues.push({ code: "payload.shapeCount-range", message: "payload.shapeCount must be zero or greater." });
  }

  if (!Array.isArray(payload.preservedFacts)) {
    issues.push({ code: "payload.preservedFacts-type", message: "payload.preservedFacts must be an array." });
  } else if (payload.preservedFacts.some((item) => typeof item !== "string")) {
    issues.push({ code: "payload.preservedFacts-item-type", message: "payload.preservedFacts items must be strings." });
  }

  validateString(payload.rationale, "payload.rationale", issues);

  return issues;
}

function normalizeBuilderPayload(payload, promptText) {
  return {
    archetype: compactText(payload.archetype || "unknown", 120) || "unknown",
    palette: compactText(payload.palette || "unknown", 120) || "unknown",
    shapeCount: Number.isInteger(payload.shapeCount) ? payload.shapeCount : null,
    preservedFacts: normalizeStringArray(payload.preservedFacts, 5, 180),
    rationale: compactText(payload.rationale || `Built slide for: ${promptText}`, 240),
  };
}

function validateJudgePayload(payload) {
  const issues = [];

  if (!isPlainObject(payload)) {
    return [{ code: "judge-payload-type", message: "judge payload must be an object." }];
  }

  validateString(payload.prompt, "payload.prompt", issues);
  validateString(payload.archetypeUsed, "payload.archetypeUsed", issues);
  validateString(payload.archetypeExpected, "payload.archetypeExpected", issues);
  validateString(payload.summary, "payload.summary", issues);

  if (!isPlainObject(payload.dimensionScores)) {
    issues.push({ code: "payload.dimensionScores-type", message: "payload.dimensionScores must be an object." });
  } else {
    for (const key of JUDGE_DIMENSION_KEYS) {
      const dimension = payload.dimensionScores[key];
      if (!isPlainObject(dimension)) {
        issues.push({ code: `payload.dimensionScores.${key}-type`, message: `payload.dimensionScores.${key} must be an object.` });
        continue;
      }

      validateInteger(dimension.score, `payload.dimensionScores.${key}.score`, issues, { min: 0, max: 2 });
      validateString(dimension.reason, `payload.dimensionScores.${key}.reason`, issues);
    }
  }

  validateInteger(payload.totalScore, "payload.totalScore", issues, { min: 0, max: JUDGE_DIMENSION_KEYS.length * 2 });
  validateInteger(payload.maxScore, "payload.maxScore", issues, { min: 0, max: JUDGE_DIMENSION_KEYS.length * 2 });

  if (!Array.isArray(payload.gaps)) {
    issues.push({ code: "payload.gaps-type", message: "payload.gaps must be an array." });
  } else if (payload.gaps.some((item) => typeof item !== "string")) {
    issues.push({ code: "payload.gaps-item-type", message: "payload.gaps items must be strings." });
  }

  if (Number.isInteger(payload.maxScore) && payload.maxScore !== JUDGE_DIMENSION_KEYS.length * 2) {
    issues.push({
      code: "payload.maxScore-contract",
      message: `payload.maxScore must equal ${JUDGE_DIMENSION_KEYS.length * 2}.`,
      semantic: true,
    });
  }

  if (isPlainObject(payload.dimensionScores)) {
    const availableScores = JUDGE_DIMENSION_KEYS
      .map((key) => payload.dimensionScores[key]?.score)
      .filter(Number.isInteger);

    if (availableScores.length === JUDGE_DIMENSION_KEYS.length && Number.isInteger(payload.totalScore)) {
      const computedTotal = availableScores.reduce((sum, score) => sum + score, 0);
      if (computedTotal !== payload.totalScore) {
        issues.push({
          code: "payload.totalScore-mismatch",
          message: `payload.totalScore (${payload.totalScore}) must equal the sum of dimension scores (${computedTotal}).`,
          semantic: true,
        });
      }
    }
  }

  return issues;
}

function normalizeJudgePayload(payload) {
  const dimensionScores = {};
  for (const key of JUDGE_DIMENSION_KEYS) {
    const dimension = payload.dimensionScores?.[key] || {};
    dimensionScores[key] = {
      score: Number.isInteger(dimension.score) ? dimension.score : 0,
      reason: compactText(dimension.reason || "", 240),
    };
  }

  return {
    prompt: compactText(payload.prompt, 240),
    archetypeUsed: compactText(payload.archetypeUsed || "unknown", 120) || "unknown",
    archetypeExpected: compactText(payload.archetypeExpected || "unknown", 120) || "unknown",
    summary: compactText(payload.summary || "", 240),
    dimensionScores,
    totalScore: Number.isInteger(payload.totalScore) ? payload.totalScore : null,
    maxScore: Number.isInteger(payload.maxScore) ? payload.maxScore : null,
    gaps: normalizeStringArray(payload.gaps, 8, 180),
  };
}

function parseScoreFromText(text) {
  const match = text.match(/(?:Total|TOTAL|Score)[:\s]*\*?\*?(\d+)\s*\/\s*(\d+)/i)
    || text.match(/\*?\*?(\d+)\s*\/\s*(\d+)\s*(?:—|–|-)/);

  return match
    ? { score: Number.parseInt(match[1], 10), maxScore: Number.parseInt(match[2], 10) }
    : { score: null, maxScore: null };
}

function parseGapsFromText(text) {
  const match = text.match(/(?:GAPS?|Primary gap|issues?|fix|missing|improve)[:\s]*([\s\S]*?)(?:\n\n|\n##|$)/i);
  if (!match?.[1]) {
    return [];
  }

  return normalizeStringArray(
    match[1]
      .split(/[;\n]/)
      .map((item) => item.replace(/^[\-\d.\s]+/, "")),
    8,
    180
  );
}

function parseLegacyBuilderObject(value, promptText, failure) {
  const issues = validateBuilderPayload(value);
  if (issues.length > 0) {
    return {
      ok: true,
      value: {
        archetype: "unknown",
        palette: "unknown",
        shapeCount: null,
        preservedFacts: [],
        rationale: compactText(value?.rationale || value?.summary || `Built slide for: ${promptText}`, 240),
      },
      validation: {
        status: "fallback",
        contract: CONTRACTS.builderSummaryResponse,
        source: "text-fallback",
        failure: {
          ...failure,
          issues,
        },
      },
    };
  }

  return {
    ok: true,
    value: normalizeBuilderPayload(value, promptText),
    validation: {
      status: "fallback",
      contract: CONTRACTS.builderSummaryResponse,
      source: "legacy-json",
      failure: {
        ...failure,
        issues,
      },
    },
  };
}

function classifyJudgeIssues(issues) {
  return issues.some((issue) => issue.semantic)
    ? FAILURE_CATEGORIES.reviewInvalid
    : FAILURE_CATEGORIES.schemaError;
}

function parseLegacyJudgeObject(value, failure) {
  const issues = validateJudgePayload(value);
  if (issues.length > 0) {
    const { score, maxScore } = parseScoreFromText(JSON.stringify(value));
    if (score == null || maxScore == null) {
      return {
        ok: false,
        failure: {
          ...failure,
          category: classifyJudgeIssues(issues),
          issues,
        },
      };
    }

    return {
      ok: true,
      value: {
        prompt: compactText(value?.prompt || "", 240),
        archetypeUsed: compactText(value?.archetypeUsed || "unknown", 120) || "unknown",
        archetypeExpected: compactText(value?.archetypeExpected || "unknown", 120) || "unknown",
        summary: compactText(value?.summary || "", 240),
        dimensionScores: {},
        totalScore: score,
        maxScore,
        gaps: normalizeStringArray(value?.gaps, 8, 180),
      },
      validation: {
        status: "fallback",
        contract: CONTRACTS.judgeResponse,
        source: "legacy-json-partial",
        failure: {
          ...failure,
          category: classifyJudgeIssues(issues),
          issues,
        },
      },
    };
  }

  return {
    ok: true,
    value: normalizeJudgePayload(value),
    validation: {
      status: "fallback",
      contract: CONTRACTS.judgeResponse,
      source: "legacy-json",
      failure: {
        ...failure,
        issues,
      },
    },
  };
}

function parseTextFallbackJudge(raw, failure) {
  const { score, maxScore } = parseScoreFromText(raw);
  if (score == null || maxScore == null) {
    return {
      ok: false,
      failure,
    };
  }

  return {
    ok: true,
    value: {
      prompt: "",
      archetypeUsed: "unknown",
      archetypeExpected: "unknown",
      summary: compactText(raw, 240),
      dimensionScores: {},
      totalScore: score,
      maxScore,
      gaps: parseGapsFromText(raw),
    },
    validation: {
      status: "fallback",
      contract: CONTRACTS.judgeResponse,
      source: "text-fallback",
      failure,
    },
  };
}

export function parseBuilderSummaryResponse(raw, { promptText = "", allowLegacyFallback = true } = {}) {
  const parsed = extractJsonObject(raw);
  if (!parsed.ok) {
    const failure = createProtocolFailure(
      FAILURE_CATEGORIES.schemaError,
      `Builder summary did not match ${CONTRACTS.builderSummaryResponse}: ${parsed.error}`,
      { contract: CONTRACTS.builderSummaryResponse, rawSnippet: compactText(raw, 240) }
    );

    if (!allowLegacyFallback) {
      return { ok: false, failure };
    }

    return {
      ok: true,
      value: {
        archetype: "unknown",
        palette: "unknown",
        shapeCount: null,
        preservedFacts: [],
        rationale: compactText(raw || `Built slide for: ${promptText}`, 240),
      },
      validation: {
        status: "fallback",
        contract: CONTRACTS.builderSummaryResponse,
        source: "text-fallback",
        failure,
      },
    };
  }

  if (parsed.value.contract === CONTRACTS.builderSummaryResponse) {
    const issues = validateBuilderPayload(parsed.value.payload);
    if (issues.length > 0) {
      return {
        ok: false,
        failure: createProtocolFailure(
          FAILURE_CATEGORIES.schemaError,
          `Builder summary payload failed validation for ${CONTRACTS.builderSummaryResponse}.`,
          {
            contract: CONTRACTS.builderSummaryResponse,
            issues,
            rawSnippet: compactText(parsed.jsonText, 240),
          }
        ),
      };
    }

    return {
      ok: true,
      value: normalizeBuilderPayload(parsed.value.payload, promptText),
      validation: {
        status: "strict",
        contract: CONTRACTS.builderSummaryResponse,
        source: parsed.source,
      },
    };
  }

  if (!allowLegacyFallback) {
    return {
      ok: false,
      failure: createProtocolFailure(
        FAILURE_CATEGORIES.schemaError,
        `Builder summary contract must be ${CONTRACTS.builderSummaryResponse}.`,
        {
          contract: CONTRACTS.builderSummaryResponse,
          receivedContract: parsed.value.contract || null,
          rawSnippet: compactText(parsed.jsonText, 240),
        }
      ),
    };
  }

  const legacyBuilderCandidate = isPlainObject(parsed.value.payload)
    ? parsed.value.payload
    : parsed.value;

  return parseLegacyBuilderObject(
    legacyBuilderCandidate,
    promptText,
    createProtocolFailure(
      FAILURE_CATEGORIES.schemaError,
      `Builder summary fell back to legacy parsing because contract ${CONTRACTS.builderSummaryResponse} was missing.`,
      {
        contract: CONTRACTS.builderSummaryResponse,
        receivedContract: parsed.value.contract || null,
        rawSnippet: compactText(parsed.jsonText, 240),
      }
    )
  );
}

export function parseJudgeResponse(raw, { allowLegacyFallback = true } = {}) {
  const parsed = extractJsonObject(raw);
  if (!parsed.ok) {
    const failure = createProtocolFailure(
      FAILURE_CATEGORIES.schemaError,
      `Judge response did not match ${CONTRACTS.judgeResponse}: ${parsed.error}`,
      { contract: CONTRACTS.judgeResponse, rawSnippet: compactText(raw, 240) }
    );

    if (!allowLegacyFallback) {
      return { ok: false, failure };
    }

    return parseTextFallbackJudge(raw, failure);
  }

  if (parsed.value.contract === CONTRACTS.judgeResponse) {
    const issues = validateJudgePayload(parsed.value.payload);
    if (issues.length > 0) {
      const failure = createProtocolFailure(
        classifyJudgeIssues(issues),
        `Judge payload failed validation for ${CONTRACTS.judgeResponse}.`,
        {
          contract: CONTRACTS.judgeResponse,
          issues,
          rawSnippet: compactText(parsed.jsonText, 240),
        }
      );

      if (!allowLegacyFallback) {
        return { ok: false, failure };
      }

      return parseTextFallbackJudge(raw, failure);
    }

    return {
      ok: true,
      value: normalizeJudgePayload(parsed.value.payload),
      validation: {
        status: "strict",
        contract: CONTRACTS.judgeResponse,
        source: parsed.source,
      },
    };
  }

  if (!allowLegacyFallback) {
    return {
      ok: false,
      failure: createProtocolFailure(
        FAILURE_CATEGORIES.schemaError,
        `Judge response contract must be ${CONTRACTS.judgeResponse}.`,
        {
          contract: CONTRACTS.judgeResponse,
          receivedContract: parsed.value.contract || null,
          rawSnippet: compactText(parsed.jsonText, 240),
        }
      ),
    };
  }

  const legacyJudgeCandidate = isPlainObject(parsed.value.payload)
    ? parsed.value.payload
    : parsed.value;

  return parseLegacyJudgeObject(
    legacyJudgeCandidate,
    createProtocolFailure(
      FAILURE_CATEGORIES.schemaError,
      `Judge response fell back to legacy parsing because contract ${CONTRACTS.judgeResponse} was missing.`,
      {
        contract: CONTRACTS.judgeResponse,
        receivedContract: parsed.value.contract || null,
        rawSnippet: compactText(parsed.jsonText, 240),
      }
    )
  );
}
