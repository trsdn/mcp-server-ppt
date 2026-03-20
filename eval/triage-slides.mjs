import {
  appendFileSync,
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  renameSync,
  unlinkSync,
  writeFileSync,
} from "fs";
import { basename, dirname, join, resolve } from "path";
import { pathToFileURL } from "url";
import {
  createFailureDetails,
  FAILURE_CATEGORIES,
  formatProtocolExample,
  getJudgeResponseSchemaExample,
  JUDGE_DIMENSION_KEYS,
  parseJudgeResponse,
} from "./lib/protocol/index.mjs";
import {
  createRuntime,
  destroyRuntime,
  executeRuntimeRequest,
  executeWithRetry,
  loadInstructionsFile,
  verifyArtifactFile,
} from "./lib/runtime/index.mjs";
import {
  EVAL_ASSET_REPO_ROOT_ENVIRONMENT_VARIABLE,
  EVAL_INPUT_ROOT,
  EVAL_OUTPUT_ROOT,
  EVAL_ROOT,
} from "./lib/runtime/environment.mjs";

const DEFAULT_MODEL = "claude-sonnet-4.5";
const DEFAULT_SOURCE_DIR = join(EVAL_INPUT_ROOT, "individual-slides");
const DEFAULT_OUTPUT_ROOT = join(EVAL_OUTPUT_ROOT, "slide-triage");
const DEFAULT_GOOD_DIR = join(DEFAULT_OUTPUT_ROOT, "good");
const DEFAULT_REJECT_DIR = join(DEFAULT_OUTPUT_ROOT, "reject");
const DEFAULT_RESULTS_PATH = join(DEFAULT_OUTPUT_ROOT, "triage-results.jsonl");
const DEFAULT_SUMMARY_PATH = join(DEFAULT_OUTPUT_ROOT, "summary.json");
const DEFAULT_INSTRUCTIONS_FILE = "agents\\judge-instructions.md";
const DEFAULT_TIMEOUT_MS = 120000;
const DEFAULT_SESSION_BATCH_SIZE = 25;
const DEFAULT_MIN_SCORE = 16;
const RATE_LIMIT_FALLBACK_DELAY_MS = 15 * 60 * 1000;

function printHelp() {
  console.log(`Direct-LLM slide triage for PNG slides under the configured eval asset root.

Usage:
  node eval\\triage-slides.mjs [options]

Options:
  --source <dir>               Source directory with PNG slides
  --output-root <dir>          Root output directory for triage artifacts
  --good-dir <dir>             Accepted slide target directory
  --reject-dir <dir>           Rejected slide target directory
  --results <file>             JSONL output with per-slide judgments
  --summary <file>             Summary JSON output
  --instructions-file <file>   Judge instructions file relative to eval\\
  --model <name>               Copilot SDK model to use (default: ${DEFAULT_MODEL})
  --reasoning-effort <level>   Optional reasoning effort for the judge session
  --timeout-ms <ms>            Per-slide timeout (default: ${DEFAULT_TIMEOUT_MS})
  --session-batch-size <n>     Slides per reused judge session (default: ${DEFAULT_SESSION_BATCH_SIZE})
  --min-score <n>              Minimum total score for acceptance (default: ${DEFAULT_MIN_SCORE})
  --shard-count <n>            Total worker shards for deterministic partitioning
  --shard-index <n>            Zero-based shard index for this worker
  --limit <n>                  Only process the first n PNGs
  --copy                       Copy files instead of moving them
  --dry-run                    Score slides but do not move/copy files
  --help                       Show this help

Defaults:
  Source:  ${DEFAULT_SOURCE_DIR}
  Good:    ${DEFAULT_GOOD_DIR}
  Reject:  ${DEFAULT_REJECT_DIR}

Asset root override:
  Set ${EVAL_ASSET_REPO_ROOT_ENVIRONMENT_VARIABLE} to a private repo clone root.
  The harness expects eval\\input, eval\\output, eval\\results, and eval\\data beneath that repo.

Selection policy:
  A slide is accepted only if it is consultant-grade and visually clean:
  - totalScore >= min-score
  - every key dimension is at least 1
  - actionTitle may be 0 for title/admin slides where topic labels are normal
  - visualExecution must be 2
  - infoHierarchy must be at least 2
`);
}

function parseArgs(argv) {
  const options = {
    sourceDir: DEFAULT_SOURCE_DIR,
    outputRoot: DEFAULT_OUTPUT_ROOT,
    goodDir: DEFAULT_GOOD_DIR,
    rejectDir: DEFAULT_REJECT_DIR,
    resultsPath: DEFAULT_RESULTS_PATH,
    summaryPath: DEFAULT_SUMMARY_PATH,
    instructionsFile: DEFAULT_INSTRUCTIONS_FILE,
    model: DEFAULT_MODEL,
    reasoningEffort: undefined,
    timeoutMs: DEFAULT_TIMEOUT_MS,
    sessionBatchSize: DEFAULT_SESSION_BATCH_SIZE,
    minScore: DEFAULT_MIN_SCORE,
    shardCount: 1,
    shardIndex: 0,
    limit: Infinity,
    copyFiles: false,
    dryRun: false,
  };
  const explicit = {
    goodDir: false,
    rejectDir: false,
    resultsPath: false,
    summaryPath: false,
  };

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    const nextValue = () => {
      const value = argv[i + 1];
      if (!value || value.startsWith("--")) {
        throw new Error(`Missing value for ${arg}`);
      }
      i += 1;
      return value;
    };

    switch (arg) {
      case "--source":
        options.sourceDir = resolve(nextValue());
        break;
      case "--output-root":
        options.outputRoot = resolve(nextValue());
        break;
      case "--good-dir":
        options.goodDir = resolve(nextValue());
        explicit.goodDir = true;
        break;
      case "--reject-dir":
        options.rejectDir = resolve(nextValue());
        explicit.rejectDir = true;
        break;
      case "--results":
        options.resultsPath = resolve(nextValue());
        explicit.resultsPath = true;
        break;
      case "--summary":
        options.summaryPath = resolve(nextValue());
        explicit.summaryPath = true;
        break;
      case "--instructions-file":
        options.instructionsFile = nextValue();
        break;
      case "--model":
        options.model = nextValue();
        break;
      case "--reasoning-effort":
        options.reasoningEffort = nextValue();
        break;
      case "--timeout-ms":
        options.timeoutMs = Number.parseInt(nextValue(), 10);
        break;
      case "--session-batch-size":
        options.sessionBatchSize = Number.parseInt(nextValue(), 10);
        break;
      case "--min-score":
        options.minScore = Number.parseInt(nextValue(), 10);
        break;
      case "--shard-count":
        options.shardCount = Number.parseInt(nextValue(), 10);
        break;
      case "--shard-index":
        options.shardIndex = Number.parseInt(nextValue(), 10);
        break;
      case "--limit":
        options.limit = Number.parseInt(nextValue(), 10);
        break;
      case "--copy":
        options.copyFiles = true;
        break;
      case "--dry-run":
        options.dryRun = true;
        break;
      case "--help":
        printHelp();
        process.exit(0);
        break;
      default:
        throw new Error(`Unknown argument: ${arg}`);
    }
  }

  if (!Number.isFinite(options.timeoutMs) || options.timeoutMs <= 0) {
    throw new Error("timeoutMs must be a positive integer");
  }

  if (!Number.isFinite(options.sessionBatchSize) || options.sessionBatchSize <= 0) {
    throw new Error("sessionBatchSize must be a positive integer");
  }

  if (!Number.isFinite(options.minScore) || options.minScore < 0) {
    throw new Error("minScore must be a non-negative integer");
  }

  if (!Number.isFinite(options.shardCount) || options.shardCount <= 0) {
    throw new Error("shardCount must be a positive integer");
  }

  if (!Number.isFinite(options.shardIndex) || options.shardIndex < 0 || options.shardIndex >= options.shardCount) {
    throw new Error("shardIndex must be between 0 and shardCount - 1");
  }

  if (options.limit !== Infinity && (!Number.isFinite(options.limit) || options.limit <= 0)) {
    throw new Error("limit must be a positive integer");
  }

  if (!explicit.goodDir && options.goodDir === DEFAULT_GOOD_DIR && options.outputRoot !== DEFAULT_OUTPUT_ROOT) {
    options.goodDir = join(options.outputRoot, "good");
  } else if (!options.goodDir) {
    options.goodDir = join(options.outputRoot, "good");
  }

  if (!explicit.rejectDir && options.rejectDir === DEFAULT_REJECT_DIR && options.outputRoot !== DEFAULT_OUTPUT_ROOT) {
    options.rejectDir = join(options.outputRoot, "reject");
  } else if (!options.rejectDir) {
    options.rejectDir = join(options.outputRoot, "reject");
  }

  if (!explicit.resultsPath && options.resultsPath === DEFAULT_RESULTS_PATH && options.outputRoot !== DEFAULT_OUTPUT_ROOT) {
    options.resultsPath = join(options.outputRoot, "triage-results.jsonl");
  } else if (!options.resultsPath) {
    options.resultsPath = join(options.outputRoot, "triage-results.jsonl");
  }

  if (!explicit.summaryPath && options.summaryPath === DEFAULT_SUMMARY_PATH && options.outputRoot !== DEFAULT_OUTPUT_ROOT) {
    options.summaryPath = join(options.outputRoot, "summary.json");
  } else if (!options.summaryPath) {
    options.summaryPath = join(options.outputRoot, "summary.json");
  }

  return options;
}

function ensureDir(path) {
  mkdirSync(path, { recursive: true });
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function prepareSlideForJudge(sourcePath, stagingDir, index) {
  ensureDir(stagingDir);
  const stagedPath = join(stagingDir, `slide-${String(index + 1).padStart(5, "0")}.png`);
  copyFileSync(sourcePath, stagedPath);
  return stagedPath;
}

function cleanupStagedSlide(stagedPath) {
  if (stagedPath && existsSync(stagedPath)) {
    unlinkSync(stagedPath);
  }
}

function listPngFiles(sourceDir, limit) {
  const files = readdirSync(sourceDir, { withFileTypes: true })
    .filter((entry) => entry.isFile() && entry.name.toLowerCase().endsWith(".png"))
    .map((entry) => join(sourceDir, entry.name))
    .sort((a, b) => basename(a).localeCompare(basename(b)));

  return Number.isFinite(limit) ? files.slice(0, limit) : files;
}

function hashFileName(fileName) {
  let hash = 2166136261;
  for (let index = 0; index < fileName.length; index++) {
    hash ^= fileName.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}

function matchesShard(filePath, shardCount, shardIndex) {
  if (shardCount <= 1) {
    return true;
  }
  return hashFileName(basename(filePath)) % shardCount === shardIndex;
}

function loadExistingResults(resultsPath) {
  if (!existsSync(resultsPath)) {
    return [];
  }

  return readFileSync(resultsPath, "utf-8")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .flatMap((line) => {
      try {
        return [JSON.parse(line)];
      } catch {
        return [];
      }
    });
}

function getProcessedSourceNames(existingResults) {
  return new Set(
    existingResults
      .map((entry) => entry?.sourceName)
      .filter(Boolean)
  );
}

function buildJudgeInitPrompt(judgeInstructions) {
  return `${judgeInstructions}

You are running a strict reference-slide triage pass over existing slide PNG files.
There is no original build prompt. Infer the intended message, archetype, and quality from the image itself.
Be conservative:
- score 2 only for clearly strong, reusable consultant-grade execution
- use 1 for anything merely acceptable
- use 0 for clear structural or visual weaknesses
- if unsure between accept and reject, reject

Always return JSON objects only on review turns. Acknowledge this instruction with:
{"acknowledged": true}`;
}

function buildJudgePrompt(pngPath) {
  const fileUri = pathToFileURL(pngPath).href;
  return `STRICT TRIAGE TASK

PNG_PATH: ${pngPath}
FILE_URI: ${fileUri}
FILE_NAME: ${basename(pngPath)}

This is an existing reference slide. There is no original prompt.
Infer the likely purpose and best-fit archetype from the visible slide itself.

Scoring intent for this curation pass:
- only consultant-grade reference slides should receive high scores
- visualExecution must reflect actual readability, spacing, overlap, alignment, and polish
- if a slide is merely decent, score it as decent rather than great
- use the full 0-2 scale honestly
- if the Windows path is awkward to open, use FILE_URI instead

Populate the response as follows:
- prompt: a concise inferred description of the slide's purpose
- archetypeUsed: the structure the slide actually uses
- archetypeExpected: the structure you believe best fits the visible content

Return JSON only. No prose, no markdown.

RESPONSE CONTRACT:
${formatProtocolExample(getJudgeResponseSchemaExample())}`;
}

function isTechnicalEvaluationFailure(payload) {
  const haystack = [
    payload?.prompt,
    payload?.summary,
    ...(Array.isArray(payload?.gaps) ? payload.gaps : []),
    ...JUDGE_DIMENSION_KEYS.map((key) => payload?.dimensionScores?.[key]?.reason),
  ]
    .filter(Boolean)
    .join(" ")
    .toLowerCase();

  return [
    "unable to evaluate",
    "could not be loaded",
    "timed out during load",
    "technical failure",
    "unable to assess",
    "image file access timed out",
    "file could not be reviewed",
  ].some((needle) => haystack.includes(needle));
}

function extractRateLimitDelayMs(message) {
  const text = String(message || "");
  const lower = text.toLowerCase();
  if (!lower.includes("rate limit") && !lower.includes("429")) {
    return null;
  }

  const minuteMatch = lower.match(/try again in\s+(\d+)\s+minute/);
  if (minuteMatch) {
    return (Number.parseInt(minuteMatch[1], 10) * 60 * 1000) + 30000;
  }

  const secondMatch = lower.match(/try again in\s+(\d+)\s+second/);
  if (secondMatch) {
    return (Number.parseInt(secondMatch[1], 10) * 1000) + 5000;
  }

  return RATE_LIMIT_FALLBACK_DELAY_MS;
}

function isActionTitleOptional(payload) {
  const haystack = [
    payload?.archetypeUsed,
    payload?.archetypeExpected,
    payload?.prompt,
    payload?.summary,
  ]
    .filter(Boolean)
    .join(" ")
    .toLowerCase();

  return [
    "title",
    "cover",
    "section",
    "divider",
    "agenda",
    "disclaimer",
    "legal",
    "confidential",
    "thank you",
    "thank-you",
    "closing",
    "appendix",
  ].some((needle) => haystack.includes(needle));
}

function shouldAccept(payload, minScore) {
  const dimensionScores = payload?.dimensionScores || {};
  const requiredDimensionKeys = JUDGE_DIMENSION_KEYS.filter(
    (key) => !(key === "actionTitle" && isActionTitleOptional(payload))
  );
  const requiredDimensionScores = requiredDimensionKeys.map((key) =>
    Number(dimensionScores?.[key]?.score ?? 0)
  );

  const noZeros = requiredDimensionScores.every((score) => score >= 1);
  const visualExecution = Number(dimensionScores?.visualExecution?.score ?? 0);
  const infoHierarchy = Number(dimensionScores?.infoHierarchy?.score ?? 0);

  return Number(payload?.totalScore ?? 0) >= minScore
    && noZeros
    && visualExecution === 2
    && infoHierarchy >= 2;
}

function writeSummary(summaryPath, summary) {
  writeFileSync(summaryPath, `${JSON.stringify(summary, null, 2)}\n`, "utf-8");
}

function moveOrCopyFile(sourcePath, targetPath, copyFiles) {
  if (existsSync(targetPath)) {
    throw new Error(`Target already exists: ${targetPath}`);
  }

  if (copyFiles) {
    copyFileSync(sourcePath, targetPath);
    return "copied";
  }

  renameSync(sourcePath, targetPath);
  return "moved";
}

async function startJudgeRuntime(agent, judgeInstructions, timeoutMs) {
  const runtime = await createRuntime(agent, { executionMode: "reuse-session" });
  const initResponse = await executeRuntimeRequest(runtime, {
    type: "prompt",
    prompt: buildJudgeInitPrompt(judgeInstructions),
    timeoutMs,
    workerTimeoutMs: timeoutMs + 30000,
  });

  if (!initResponse.ok) {
    const failure = createFailureDetails(initResponse.error, {
      fallbackCategory: initResponse.errorCategory || FAILURE_CATEGORIES.runtimeUnavailable,
    });
    await destroyRuntime(runtime);
    throw new Error(`Failed to initialize judge runtime: ${failure.message}`);
  }

  return runtime;
}

async function startJudgeRuntimeWithRetry(agent, judgeInstructions, timeoutMs) {
  while (true) {
    try {
      return await startJudgeRuntime(agent, judgeInstructions, timeoutMs);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      const delayMs = extractRateLimitDelayMs(message);
      if (!delayMs) {
        throw error;
      }

      console.log(`  rate limited during runtime start; waiting ${Math.ceil(delayMs / 60000)} min before retry`);
      await sleep(delayMs);
    }
  }
}

async function scoreSlide(runtime, pngPath, timeoutMs) {
  const response = await executeRuntimeRequest(runtime, {
    type: "prompt",
    prompt: buildJudgePrompt(pngPath),
    timeoutMs,
    workerTimeoutMs: timeoutMs + 30000,
  });

  if (!response.ok) {
    const failure = createFailureDetails(response.error, {
      fallbackCategory: response.errorCategory || FAILURE_CATEGORIES.toolFailure,
    });
    return {
      ok: false,
      error: failure.message,
      errorCategory: failure.category,
      failure,
      rateLimitDelayMs: extractRateLimitDelayMs(failure.message),
    };
  }

  const raw = response.content || "";
  const parsed = parseJudgeResponse(raw, { allowLegacyFallback: true });
  if (!parsed.ok) {
    const failure = createFailureDetails(parsed.failure.message, {
      fallbackCategory: parsed.failure.category,
    });
    return {
      ok: false,
      error: failure.message,
      errorCategory: failure.category,
      failure,
      raw,
    };
  }

  if (isTechnicalEvaluationFailure(parsed.value)) {
    const failure = createFailureDetails(
      "Judge returned a technical image-load failure instead of a usable slide assessment.",
      { fallbackCategory: FAILURE_CATEGORIES.runtimeUnavailable }
    );
    return {
      ok: false,
      error: failure.message,
      errorCategory: failure.category,
      failure,
      raw,
    };
  }

  return {
    ok: true,
    raw,
    payload: parsed.value,
    validation: parsed.validation,
  };
}

function createSummary({
  options,
  startedAt,
  processedCount,
  acceptedCount,
  rejectedCount,
  errorCount,
  remainingCount,
  resultsPath,
}) {
  return {
    generatedAt: new Date().toISOString(),
    startedAt,
    finishedAt: new Date().toISOString(),
    sourceDir: options.sourceDir,
    goodDir: options.goodDir,
    rejectDir: options.rejectDir,
    resultsPath,
    mode: options.copyFiles ? "copy" : "move",
    dryRun: options.dryRun,
    model: options.model,
    reasoningEffort: options.reasoningEffort || null,
    sessionBatchSize: options.sessionBatchSize,
    minScore: options.minScore,
    shardCount: options.shardCount,
    shardIndex: options.shardIndex,
    processedCount,
    acceptedCount,
    rejectedCount,
    errorCount,
    remainingCount,
  };
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const stagingDir = join(options.outputRoot, "_staging");

  if (!existsSync(options.sourceDir)) {
    throw new Error(`Source directory not found: ${options.sourceDir}`);
  }

  ensureDir(options.outputRoot);
  ensureDir(options.goodDir);
  ensureDir(options.rejectDir);
  ensureDir(dirname(options.resultsPath));
  ensureDir(dirname(options.summaryPath));

  const instructions = loadInstructionsFile({
    baseDir: EVAL_ROOT,
    instructionsFile: options.instructionsFile,
    label: "triage judge instructions",
  });

  if (!instructions.ok) {
    throw new Error(instructions.message);
  }

  const existingResults = loadExistingResults(options.resultsPath);
  const processedSourceNames = getProcessedSourceNames(existingResults);
  const candidateFiles = listPngFiles(options.sourceDir, options.limit)
    .filter((filePath) => matchesShard(filePath, options.shardCount, options.shardIndex))
    .filter((filePath) => !processedSourceNames.has(basename(filePath)));

  console.log(`Triage source: ${options.sourceDir}`);
  console.log(`Pending PNGs: ${candidateFiles.length}`);
  console.log(`Accepted -> ${options.goodDir}`);
  console.log(`Rejected -> ${options.rejectDir}`);
  console.log(`Model: ${options.model}${options.reasoningEffort ? ` (${options.reasoningEffort})` : ""}`);
  if (options.shardCount > 1) {
    console.log(`Shard: ${options.shardIndex + 1}/${options.shardCount}`);
  }
  console.log(
    `Selection: totalScore >= ${options.minScore}, no zero key dimensions, actionTitle exempt on title/admin slides, visualExecution=2, infoHierarchy>=2`
  );
  if (options.dryRun) {
    console.log("Mode: dry-run (no file move/copy)");
  } else {
    console.log(`Mode: ${options.copyFiles ? "copy" : "move"}`);
  }

  if (candidateFiles.length === 0) {
    const remaining = listPngFiles(options.sourceDir, Infinity).length;
    writeSummary(options.summaryPath, createSummary({
      options,
      startedAt: new Date().toISOString(),
      processedCount: 0,
      acceptedCount: 0,
      rejectedCount: 0,
      errorCount: 0,
      remainingCount: remaining,
      resultsPath: options.resultsPath,
    }));
    console.log("Nothing to do.");
    return;
  }

  const agent = {
    model: options.model,
    runtime: "copilot-sdk",
    executionMode: "reuse-session",
  };
  if (options.reasoningEffort) {
    agent.reasoningEffort = options.reasoningEffort;
  }

  const startedAt = new Date().toISOString();
  let runtime = null;
  let slidesInCurrentSession = 0;
  let acceptedCount = 0;
  let rejectedCount = 0;
  let errorCount = 0;
  let processedCount = 0;

  try {
    for (let index = 0; index < candidateFiles.length; index++) {
      const sourcePath = candidateFiles[index];
      const sourceName = basename(sourcePath);
      // File may have been moved by another shard/process between directory scan and now
      if (!existsSync(sourcePath)) {
        console.log(
          `[${index + 1}/${candidateFiles.length}] skipped (already moved) ${sourceName}`
        );
        continue;
      }
      const stagedPath = prepareSlideForJudge(sourcePath, stagingDir, index);

      const fileStatus = verifyArtifactFile(stagedPath, { kind: "png" });
      if (!fileStatus.ok) {
        cleanupStagedSlide(stagedPath);
        errorCount += 1;
        const errorRecord = {
          processedAt: new Date().toISOString(),
          sourceName,
          sourcePath,
          accepted: false,
          action: "error",
          error: fileStatus.message,
          errorCategory: fileStatus.category,
        };
        appendFileSync(options.resultsPath, `${JSON.stringify(errorRecord)}\n`, "utf-8");
        console.log(`[${index + 1}/${candidateFiles.length}] error ${sourceName}: ${fileStatus.message}`);
        continue;
      }

      if (!runtime || slidesInCurrentSession >= options.sessionBatchSize) {
        if (runtime) {
          await destroyRuntime(runtime);
        }
        runtime = await startJudgeRuntimeWithRetry(agent, instructions.text, options.timeoutMs);
        slidesInCurrentSession = 0;
      }

      let scoreResult;
      try {
        scoreResult = await executeWithRetry(
          async () => scoreSlide(runtime, stagedPath, options.timeoutMs),
          {
            maxAttempts: 2,
            baseDelayMs: 2000,
            shouldRetry: ({ result, failure }) => {
              const message = String(result?.error || failure?.message || "");
              return message.includes("Session not found")
                || Boolean(result?.rateLimitDelayMs)
                || failure?.category === FAILURE_CATEGORIES.timeout
                || failure?.category === FAILURE_CATEGORIES.runtimeUnavailable
                || failure?.category === FAILURE_CATEGORIES.schemaError;
            },
            onRetry: async ({ result, failure, delayMs, attempt, maxAttempts }) => {
              const message = String(result?.error || failure?.message || "");
              console.log(`  retry ${attempt}/${maxAttempts} after ${delayMs}ms: ${message || "transient failure"}`);
              if (result?.rateLimitDelayMs) {
                console.log(`  rate limited during scoring; waiting ${Math.ceil(result.rateLimitDelayMs / 60000)} min before retry`);
                await sleep(result.rateLimitDelayMs);
              }
              if (
                message.includes("Session not found")
                || Boolean(result?.rateLimitDelayMs)
                || failure?.category === FAILURE_CATEGORIES.timeout
                || failure?.category === FAILURE_CATEGORIES.runtimeUnavailable
                || failure?.category === FAILURE_CATEGORIES.schemaError
              ) {
                if (runtime) {
                  await destroyRuntime(runtime);
                }
                runtime = await startJudgeRuntimeWithRetry(agent, instructions.text, options.timeoutMs);
                slidesInCurrentSession = 0;
              }
            },
          }
        );
      } finally {
        cleanupStagedSlide(stagedPath);
      }

      slidesInCurrentSession += 1;

      if (!scoreResult.ok) {
        errorCount += 1;
        const errorRecord = {
          processedAt: new Date().toISOString(),
          sourceName,
          sourcePath,
          accepted: false,
          action: "error",
          error: scoreResult.error,
          errorCategory: scoreResult.errorCategory || scoreResult.failure?.category || null,
          retry: scoreResult.retry || null,
        };
        appendFileSync(options.resultsPath, `${JSON.stringify(errorRecord)}\n`, "utf-8");
        console.log(`[${index + 1}/${candidateFiles.length}] error ${sourceName}: ${scoreResult.error}`);
        continue;
      }

      const accepted = shouldAccept(scoreResult.payload, options.minScore);
      const targetDir = accepted ? options.goodDir : options.rejectDir;
      const targetPath = join(targetDir, sourceName);
      let action = "scored";

      if (!options.dryRun) {
        try {
          action = moveOrCopyFile(sourcePath, targetPath, options.copyFiles);
        } catch (moveErr) {
          if (moveErr.code === "ENOENT" && !existsSync(sourcePath)) {
            // File was moved by another shard/process — skip gracefully
            console.log(
              `[${index + 1}/${candidateFiles.length}] skipped (already moved) ${sourceName}`
            );
            continue;
          }
          throw moveErr;
        }
      }

      if (accepted) {
        acceptedCount += 1;
      } else {
        rejectedCount += 1;
      }
      processedCount += 1;

      const resultRecord = {
        processedAt: new Date().toISOString(),
        sourceName,
        sourcePath,
        targetPath: options.dryRun ? null : targetPath,
        accepted,
        action,
        totalScore: scoreResult.payload.totalScore,
        maxScore: scoreResult.payload.maxScore,
        prompt: scoreResult.payload.prompt,
        archetypeUsed: scoreResult.payload.archetypeUsed,
        archetypeExpected: scoreResult.payload.archetypeExpected,
        summary: scoreResult.payload.summary,
        dimensionScores: scoreResult.payload.dimensionScores,
        gaps: scoreResult.payload.gaps,
        validation: scoreResult.validation || null,
        retry: scoreResult.retry || null,
      };
      appendFileSync(options.resultsPath, `${JSON.stringify(resultRecord)}\n`, "utf-8");

      const verdict = accepted ? "good" : "reject";
      console.log(
        `[${index + 1}/${candidateFiles.length}] ${verdict} ${scoreResult.payload.totalScore}/${scoreResult.payload.maxScore} ${sourceName}`
      );

      const remaining = options.dryRun
        ? candidateFiles.length - (index + 1)
        : listPngFiles(options.sourceDir, Infinity).length;
      writeSummary(options.summaryPath, createSummary({
        options,
        startedAt,
        processedCount,
        acceptedCount,
        rejectedCount,
        errorCount,
        remainingCount: remaining,
        resultsPath: options.resultsPath,
      }));
    }
  } finally {
    if (runtime) {
      await destroyRuntime(runtime);
    }
  }

  const remainingCount = listPngFiles(options.sourceDir, Infinity).length;
  const finalSummary = createSummary({
    options,
    startedAt,
    processedCount,
    acceptedCount,
    rejectedCount,
    errorCount,
    remainingCount,
    resultsPath: options.resultsPath,
  });
  writeSummary(options.summaryPath, finalSummary);

  console.log("");
  console.log("Done.");
  console.log(`Processed: ${processedCount}`);
  console.log(`Accepted:  ${acceptedCount}`);
  console.log(`Rejected:  ${rejectedCount}`);
  console.log(`Errors:    ${errorCount}`);
  console.log(`Remaining: ${remainingCount}`);
  console.log(`Results:   ${options.resultsPath}`);
  console.log(`Summary:   ${options.summaryPath}`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
