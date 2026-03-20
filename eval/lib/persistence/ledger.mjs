import { appendFile, mkdir, readFile } from "fs/promises";
import { join, relative, resolve } from "path";
import { createArtifactManifest, captureSkillSnapshot } from "./manifests.mjs";
import {
  createLoopRecord,
  createRunManifest,
  createTransactionEntry,
  readStructuredRecord,
  writeStructuredRecord,
} from "./records.mjs";

function sanitizePathSegment(value, fallback = "unknown") {
  const sanitized = String(value ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80);

  return sanitized || fallback;
}

function toRunStamp(value = new Date()) {
  return value.toISOString().slice(0, 19).replace(/[T:]/g, "-");
}

function buildLoopFileName(loopNumber, promptId) {
  const loopLabel = String(Number.isInteger(loopNumber) ? loopNumber : 0).padStart(4, "0");
  return `loop-${loopLabel}-${sanitizePathSegment(promptId, "prompt")}.json`;
}

function deriveCounters(loopIndex) {
  const counters = {
    recordedLoops: loopIndex.length,
    completedLoops: 0,
    failedLoops: 0,
    scoredLoops: 0,
  };

  for (const loop of loopIndex) {
    if (loop.status === "completed") {
      counters.completedLoops += 1;
    } else {
      counters.failedLoops += 1;
    }

    if (Number.isFinite(loop.score)) {
      counters.scoredLoops += 1;
    }
  }

  return counters;
}

export async function createEvalLedger({
  ledgerRoot,
  runName,
  runner,
  metadata = {},
}) {
  const runId = `${sanitizePathSegment(runName || runner, "eval-run")}-${toRunStamp()}`;
  const rootDir = resolve(ledgerRoot);
  const runDir = join(rootDir, runId);
  const loopsDir = join(runDir, "loops");
  const transactionsPath = join(runDir, "transactions.jsonl");
  const manifestPath = join(runDir, "manifest.json");
  const startedAt = new Date().toISOString();
  const loopIndex = [];

  await mkdir(loopsDir, { recursive: true });

  async function persistManifest({
    status,
    finishedAt = null,
    summary = null,
    finalArtifacts = null,
  } = {}) {
    const manifest = createRunManifest({
      runId,
      runner,
      runName,
      status: status || "running",
      startedAt,
      finishedAt,
      directories: {
        rootDir: runDir,
        loopsDir,
        manifestPath,
        transactionsPath,
      },
      metadata,
      counters: deriveCounters(loopIndex),
      loopIndex,
      finalArtifacts,
      summary,
    });

    await writeStructuredRecord(manifestPath, manifest);
    return manifest;
  }

  await persistManifest({ status: "running" });

  async function writeLoopRecord({
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
    artifacts = [],
    skills = null,
    metadata: recordMetadata = {},
  }) {
    const recordId = `${runId}:loop:${String(Number.isInteger(loopNumber) ? loopNumber : 0).padStart(4, "0")}:${sanitizePathSegment(prompt?.id, "prompt")}`;
    const loopFilePath = join(loopsDir, buildLoopFileName(loopNumber, prompt?.id));
    const artifactManifest = artifacts?.contract
      ? artifacts
      : await createArtifactManifest({ artifacts, relativeTo: runDir });

    const record = createLoopRecord({
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
      artifacts: artifactManifest,
      skills,
      metadata: recordMetadata,
    });

    await writeStructuredRecord(loopFilePath, record);

    const relativeRecordPath = relative(runDir, loopFilePath);
    const transaction = createTransactionEntry({
      entryId: `${recordId}:txn`,
      runId,
      runner,
      runName,
      loopNumber,
      promptId: prompt?.id,
      status: record.outcome.status,
      score: record.outcome.score,
      maxScore: record.outcome.maxScore,
      errorCategory: record.outcome.errorCategory,
      errorMessage: record.outcome.errorMessage,
      recordPath: relativeRecordPath,
      artifactFingerprintSha256: record.artifacts?.fingerprintSha256,
      skillSnapshotSha256: record.skills?.beforeLoop?.fingerprintSha256 || record.skills?.afterLoop?.fingerprintSha256 || null,
    });

    await appendFile(transactionsPath, `${JSON.stringify(transaction)}\n`, "utf8");

    loopIndex.push({
      loopNumber: Number.isInteger(loopNumber) ? loopNumber : null,
      promptId: prompt?.id ?? null,
      status: record.outcome.status,
      score: record.outcome.score,
      maxScore: record.outcome.maxScore,
      recordPath: relativeRecordPath,
    });

    await persistManifest({ status: "running" });

    return { record, recordPath: loopFilePath, transaction };
  }

  async function finalize({
    status = "completed",
    summary = null,
    finalArtifacts = [],
    finishedAt = new Date().toISOString(),
  } = {}) {
    const artifactManifest = finalArtifacts?.contract
      ? finalArtifacts
      : await createArtifactManifest({ artifacts: finalArtifacts, relativeTo: runDir });

    return persistManifest({
      status,
      finishedAt,
      summary,
      finalArtifacts: artifactManifest,
    });
  }

  return {
    runId,
    runDir,
    loopsDir,
    manifestPath,
    transactionsPath,
    captureSkillSnapshot,
    writeLoopRecord,
    finalize,
  };
}

export async function readRunManifest(manifestPath) {
  return readStructuredRecord(manifestPath);
}

export async function readLoopRecord(recordPath) {
  return readStructuredRecord(recordPath);
}

export async function readTransactionLedger(transactionsPath) {
  const content = await readFile(transactionsPath, "utf8");
  return content
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => JSON.parse(line));
}
