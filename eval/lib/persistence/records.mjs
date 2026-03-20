import { mkdir, readFile, writeFile } from "fs/promises";
import { dirname } from "path";

export const PERSISTENCE_CONTRACTS = Object.freeze({
  runManifest: "eval-run-manifest/v1",
  loopRecord: "eval-loop-record/v1",
  transactionEntry: "eval-transaction-entry/v1",
  artifactManifest: "eval-artifact-manifest/v1",
  skillSnapshot: "eval-skill-snapshot/v1",
});

function isPlainObject(value) {
  return value != null && typeof value === "object" && !Array.isArray(value);
}

function compactText(value, maxLength = 4000) {
  return typeof value === "string"
    ? value.replace(/\r\n/g, "\n").trim().slice(0, maxLength)
    : null;
}

function normalizeStringArray(value, maxItems = 12, maxLength = 300) {
  return Array.isArray(value)
    ? value
      .map((item) => compactText(item, maxLength))
      .filter(Boolean)
      .slice(0, maxItems)
    : [];
}

function normalizeCarryoverEntries(value, maxItems = 8) {
  return Array.isArray(value)
    ? value
      .map((item) => {
        if (typeof item === "string") {
          return compactText(item, 2000);
        }

        if (isPlainObject(item)) {
          return sanitizeObject(item);
        }

        return null;
      })
      .filter(Boolean)
      .slice(0, maxItems)
    : [];
}

function sanitizeObject(value) {
  if (Array.isArray(value)) {
    return value.map(sanitizeObject);
  }

  if (isPlainObject(value)) {
    return Object.fromEntries(
      Object.entries(value)
        .filter(([, entryValue]) => entryValue !== undefined)
        .map(([key, entryValue]) => [key, sanitizeObject(entryValue)])
    );
  }

  return value;
}

function sortValue(value) {
  if (Array.isArray(value)) {
    return value.map(sortValue);
  }

  if (isPlainObject(value)) {
    return Object.fromEntries(
      Object.keys(value)
        .sort((left, right) => left.localeCompare(right))
        .map((key) => [key, sortValue(value[key])])
    );
  }

  return value;
}

function toIsoString(value) {
  if (!value) return null;
  if (typeof value === "string") return value;
  if (value instanceof Date) return value.toISOString();

  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date.toISOString();
}

function computeTotalMs(startedAt, finishedAt, totalMs) {
  if (Number.isFinite(totalMs)) return totalMs;
  const start = startedAt ? Date.parse(startedAt) : Number.NaN;
  const end = finishedAt ? Date.parse(finishedAt) : Number.NaN;
  return Number.isFinite(start) && Number.isFinite(end) ? Math.max(0, end - start) : null;
}

function normalizePhaseTiming(value) {
  if (!isPlainObject(value)) return null;

  const startedAt = toIsoString(value.startedAt);
  const finishedAt = toIsoString(value.finishedAt);

  return sanitizeObject({
    startedAt,
    finishedAt,
    durationMs: computeTotalMs(startedAt, finishedAt, value.durationMs),
  });
}

function normalizeTimings(value) {
  const startedAt = toIsoString(value?.startedAt);
  const finishedAt = toIsoString(value?.finishedAt);
  const phases = isPlainObject(value?.phases)
    ? Object.fromEntries(
      Object.entries(value.phases)
        .map(([key, phase]) => [key, normalizePhaseTiming(phase)])
        .filter(([, phase]) => phase != null)
    )
    : {};

  return sanitizeObject({
    startedAt,
    finishedAt,
    totalMs: computeTotalMs(startedAt, finishedAt, value?.totalMs),
    phases,
  });
}

function normalizeRetries(value) {
  const retries = isPlainObject(value) ? value : {};
  return sanitizeObject({
    build: Number.isInteger(retries.build) ? retries.build : 0,
    judge: Number.isInteger(retries.judge) ? retries.judge : 0,
    improvement: Number.isInteger(retries.improvement) ? retries.improvement : 0,
  });
}

function normalizePrompt(prompt = {}) {
  return sanitizeObject({
    id: compactText(prompt.id ?? "", 160) || null,
    text: compactText(prompt.text ?? prompt.prompt ?? "", 4000) || null,
    category: compactText(prompt.category ?? prompt.archetype ?? "", 160) || null,
    archetype: compactText(prompt.archetype ?? prompt.category ?? "", 160) || null,
  });
}

function normalizeBuilderRecord(builder = {}) {
  return sanitizeObject({
    request: builder.request ?? null,
    response: {
      completion: compactText(builder.completion ?? "", 400) || null,
      summary: builder.summary ?? null,
      text: compactText(builder.text ?? "", 12000),
      raw: compactText(builder.raw ?? "", 20000),
      errorMessage: compactText(builder.errorMessage ?? builder.error ?? "", 4000),
      errorCategory: compactText(builder.errorCategory ?? "", 160) || null,
    },
    validation: builder.validation ?? null,
  });
}

function normalizeJudgeRecord(judge = {}) {
  return sanitizeObject({
    request: judge.request ?? null,
    response: {
      parsed: judge.parsed ?? null,
      summary: compactText(judge.summary ?? "", 4000),
      raw: compactText(judge.raw ?? "", 20000),
      errorMessage: compactText(judge.errorMessage ?? judge.error ?? "", 4000),
      errorCategory: compactText(judge.errorCategory ?? "", 160) || null,
    },
    validation: judge.validation ?? null,
  });
}

function normalizeImprovementRecord(improvement = {}) {
  if (!improvement || improvement.attempted !== true) {
    return sanitizeObject({
      attempted: false,
      response: null,
      errorMessage: null,
    });
  }

  return sanitizeObject({
    attempted: true,
    response: compactText(improvement.response ?? "", 12000),
    errorMessage: compactText(improvement.errorMessage ?? improvement.error ?? "", 4000),
  });
}

function deriveOutcome({
  status,
  buildStatus,
  judgeStatus,
  score,
  maxScore,
  errorCategory,
  errorMessage,
}) {
  return sanitizeObject({
    status: compactText(status ?? "", 160) || "unknown",
    buildStatus: compactText(buildStatus ?? "", 160) || "unknown",
    judgeStatus: compactText(judgeStatus ?? "", 160) || "not_run",
    score: Number.isFinite(score) ? score : null,
    maxScore: Number.isFinite(maxScore) ? maxScore : null,
    errorCategory: compactText(errorCategory ?? "", 160) || null,
    errorMessage: compactText(errorMessage ?? "", 4000),
  });
}

export function stableJsonStringify(value, space = 2) {
  return JSON.stringify(sortValue(sanitizeObject(value)), null, space);
}

export async function writeStructuredRecord(filePath, value) {
  await mkdir(dirname(filePath), { recursive: true });
  await writeFile(filePath, `${stableJsonStringify(value)}\n`, "utf8");
  return filePath;
}

export async function readStructuredRecord(filePath) {
  return JSON.parse(await readFile(filePath, "utf8"));
}

export function createLoopRecord({
  recordId,
  runId,
  runner,
  runName,
  loopNumber,
  prompt,
  status,
  buildStatus,
  judgeStatus,
  score,
  maxScore,
  errorCategory,
  errorMessage,
  timings,
  retries,
  carryover,
  diagnostics,
  builder,
  judge,
  improvement,
  artifacts,
  skills,
  metadata,
  persistedAt = new Date().toISOString(),
}) {
  return sanitizeObject({
    contract: PERSISTENCE_CONTRACTS.loopRecord,
    recordId,
    runId,
    runner,
    runName,
    loop: {
      number: Number.isInteger(loopNumber) ? loopNumber : null,
      prompt: normalizePrompt(prompt),
    },
    outcome: deriveOutcome({
      status,
      buildStatus,
      judgeStatus,
      score,
      maxScore,
      errorCategory,
      errorMessage,
    }),
    timings: normalizeTimings(timings),
    retries: normalizeRetries(retries),
    carryover: {
      builder: normalizeCarryoverEntries(carryover?.builder),
      reviewer: normalizeCarryoverEntries(carryover?.reviewer),
    },
    diagnostics: diagnostics ?? {},
    builder: normalizeBuilderRecord(builder),
    judge: normalizeJudgeRecord(judge),
    improvement: normalizeImprovementRecord(improvement),
    artifacts: artifacts ?? null,
    skills: skills ?? null,
    metadata: metadata ?? {},
    persistedAt: toIsoString(persistedAt),
  });
}

export function createTransactionEntry({
  entryId,
  runId,
  runner,
  runName,
  loopNumber,
  promptId,
  status,
  score,
  maxScore,
  errorCategory,
  errorMessage,
  recordPath,
  artifactFingerprintSha256,
  skillSnapshotSha256,
  persistedAt = new Date().toISOString(),
}) {
  return sanitizeObject({
    contract: PERSISTENCE_CONTRACTS.transactionEntry,
    entryId,
    runId,
    runner,
    runName,
    loopNumber: Number.isInteger(loopNumber) ? loopNumber : null,
    promptId: compactText(promptId ?? "", 160) || null,
    status: compactText(status ?? "", 160) || "unknown",
    score: Number.isFinite(score) ? score : null,
    maxScore: Number.isFinite(maxScore) ? maxScore : null,
    errorCategory: compactText(errorCategory ?? "", 160) || null,
    errorMessage: compactText(errorMessage ?? "", 4000),
    recordPath: compactText(recordPath ?? "", 1000) || null,
    artifactFingerprintSha256: compactText(artifactFingerprintSha256 ?? "", 128) || null,
    skillSnapshotSha256: compactText(skillSnapshotSha256 ?? "", 128) || null,
    persistedAt: toIsoString(persistedAt),
  });
}

export function createRunManifest({
  runId,
  runner,
  runName,
  status,
  startedAt,
  finishedAt,
  directories,
  metadata,
  counters,
  loopIndex,
  finalArtifacts,
  summary,
  persistedAt = new Date().toISOString(),
}) {
  return sanitizeObject({
    contract: PERSISTENCE_CONTRACTS.runManifest,
    runId,
    runner,
    runName,
    status: compactText(status ?? "", 160) || "unknown",
    startedAt: toIsoString(startedAt),
    finishedAt: toIsoString(finishedAt),
    directories: directories ?? {},
    metadata: metadata ?? {},
    counters: counters ?? {},
    loopIndex: Array.isArray(loopIndex) ? loopIndex : [],
    finalArtifacts: finalArtifacts ?? null,
    summary: summary ?? null,
    persistedAt: toIsoString(persistedAt),
  });
}
