/**
 * Archetype-focused evaluation loop.
 *
 * Fresh builder client per loop. Persistent judge session per archetype cycle.
 * Optional improver agent when gaps are found.
 *
 * Usage:
 *   node run-archetype-eval.mjs configs/big-number.json
 *   node run-archetype-eval.mjs configs/all
 */

import { readFileSync, writeFileSync, mkdirSync, readdirSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import {
  createRuntime,
  destroyRuntime,
  executeAgentRequest,
  executeWithRetry,
  loadInstructionsFile,
  normalizeTransport,
  executeRuntimeRequest,
  shouldReuseSessionContext,
  shouldUseIsolatedProcess,
  verifyBuildArtifacts,
  ARCHETYPES_DIR,
  resolveArchetypeFamily,
} from "./lib/runtime/index.mjs";
import { CLI_PATH, SKILLS_DIR, resolveEvalAssetPathFromRelative } from "./lib/runtime/environment.mjs";
import {
  EVAL_MODES,
  getModeScopedDirectory,
  getModeTaggedName,
  resolveEvalMode,
} from "./lib/mode.mjs";
import {
  FAILURE_CATEGORIES,
  createFailureDetails,
  createBuilderCarryoverEntry,
  createEvaluationRequestEnvelope,
  createJudgmentRequestEnvelope,
  createReviewerCarryoverEntry,
  formatProtocolExample,
  getBuilderSummaryResponseSchemaExample,
  getJudgeResponseSchemaExample,
  parseBuilderSummaryResponse,
  parseJudgeResponse,
} from "./lib/protocol/index.mjs";
import {
  captureSkillSnapshot,
  createEvalLedger,
} from "./lib/persistence/index.mjs";
import {
  formatRunReportForConsole,
  generateRunReportFromPersistence,
} from "./lib/reporting/index.mjs";
import {
  ORCHESTRATOR_STEPS,
  ORCHESTRATOR_STEP_STATUS,
  cleanupPptSessions,
  createLoopState,
  createOrchestrationContext,
  createStepSequence,
  ensureService,
  formatRollingContext,
  getBuilderTimeoutMs,
  pushRollingItem,
  runLoopOrchestrator,
} from "./lib/orchestrator/index.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const DEFAULT_RETRY_POLICY = Object.freeze({
  maxAttempts: 2,
  baseDelayMs: 1500,
});

function formatCarryoverEnvelope(builderCarryover = [], reviewerCarryover = []) {
  return formatProtocolExample({
    builderCarryover,
    reviewerCarryover,
  });
}

function buildBuilderSummaryPrompt(promptObj, loopContext) {
  return `Return a JSON object only. No prose, no markdown.

Summarize the slide you just built for "${promptObj.text}".

If this session is being reused, treat the structured reviewer feedback below as the authoritative prior-loop context for this turn.
${formatRollingContext(loopContext, "Latest reviewer feedback carried in this conversation")}

Response contract:
${formatProtocolExample(getBuilderSummaryResponseSchemaExample())}`;
}

async function buildSlide(config, promptObj, pngPath, builderRuntime = null, context = {}) {
  const { builder } = config;
  const pptxPath = pngPath.replace(".png", ".pptx");
  const transport = normalizeTransport(builder.transport || "cli");
  const skillPaths = (builder.skillFiles || []).map((file) => join(SKILLS_DIR, file)).join("\n- ");

  // Resolve archetype family for targeted design guidance
  const familyId = resolveArchetypeFamily(config.archetype);
  const archetypeRegistryPath = join(ARCHETYPES_DIR, "registry.md");
  const archetypeFamilyPath = join(ARCHETYPES_DIR, `${familyId}.md`);
  const evidenceDesignPath = join(ARCHETYPES_DIR, "evidence-design.md");

  const instructionsFile = loadInstructionsFile({
    baseDir: __dirname,
    instructionsFile: builder.instructionsFile,
    expectedTransport: transport,
    label: "builder instructions",
  });

  if (!instructionsFile.ok) {
    const failure = createFailureDetails(instructionsFile.message, {
      fallbackCategory: instructionsFile.category,
    });
    return {
      ok: false,
      error: failure.message,
      errorCategory: failure.category,
      failure,
      requestEnvelope: null,
      pptxPath,
    };
  }

  const instructions = instructionsFile.text.replace("{CLI_PATH}", CLI_PATH);

  const transportPrompt = transport === "mcp"
    ? `OUTPUT: PPTX=${pptxPath} PNG=${pngPath}

Use PowerPoint MCP tools only. Preferred flow: file create/open → slide create Blank → shape/text operations → export slide-to-image → file close save:true.`
    : `CLI: ${CLI_PATH}
RULES: --color not --font-color. --alignment not --horizontal-alignment. No \\n in --text. Service is running.
OUTPUT: PPTX=${pptxPath} PNG=${pngPath}

Steps: read skills → session create ${pptxPath} → slide create --layout-name Blank → build shapes → export slide-to-image → session close --save.`;

  // Build variant context for the prompt
  const variantHint = promptObj.expectedVariant
    ? `\nEXPECTED VARIANT: ${promptObj.expectedVariant} (within the ${familyId} archetype family)`
    : "";

  const requestEnvelope = createEvaluationRequestEnvelope({
    promptId: promptObj.id,
    prompt: promptObj.text,
    archetype: config.archetype,
    archetypeFamily: familyId,
    expectedVariant: promptObj.expectedVariant || null,
    transport,
    pngPath,
    pptxPath,
    builderCarryover: context.builderCarryover || [],
    reviewerCarryover: context.reviewerCarryover || [],
  });

  const prompt = `${instructions}

TASK: Build ONE ${config.archetype} slide.
PROMPT: "${promptObj.text}"${variantHint}

READ these skill files for design guidance:
- ${skillPaths}

READ these archetype references for layout and variant rules:
- ${archetypeRegistryPath}
- ${archetypeFamilyPath}
- ${evidenceDesignPath}

${formatRollingContext(context.builderCarryover || [], "Builder carry-over from earlier loops")}
${formatRollingContext(context.reviewerCarryover || [], "Reviewer carry-over from earlier loops")}
Structured carry-over envelope for this turn:
${formatCarryoverEnvelope(context.builderCarryover || [], context.reviewerCarryover || [])}

REQUEST ENVELOPE:
${formatProtocolExample(requestEnvelope)}

${transportPrompt}`;

  const timeoutMs = getBuilderTimeoutMs(builder);

  const response = builderRuntime
    ? await executeRuntimeRequest(builderRuntime, {
      type: "build",
      prompt,
      pngPath,
      pptxPath,
      timeoutMs,
      workerTimeoutMs: timeoutMs + 45000,
      summaryPrompt: buildBuilderSummaryPrompt(promptObj, context.reviewerCarryover || []),
      summaryTimeoutMs: 30000,
    })
    : await executeAgentRequest(builder, {
      type: "build",
      prompt,
      pngPath,
      pptxPath,
      timeoutMs,
      workerTimeoutMs: timeoutMs + 45000,
      summaryPrompt: buildBuilderSummaryPrompt(promptObj, context.reviewerCarryover || []),
      summaryTimeoutMs: 30000,
    });

  if (!response.ok) {
    const failure = createFailureDetails(response.error, {
      fallbackCategory: FAILURE_CATEGORIES.toolFailure,
    });
    return {
      ok: false,
      error: failure.message,
      errorCategory: failure.category,
      failure,
      requestEnvelope,
      pptxPath,
    };
  }

  const builderSummaryResult = parseBuilderSummaryResponse(response.summaryContent || "", {
    promptText: promptObj.text,
    allowLegacyFallback: true,
  });

  return {
    ok: true,
    requestEnvelope,
    pptxPath,
    completion: response.completion || "completed",
    builderSummary: builderSummaryResult.ok ? builderSummaryResult.value : null,
    builderValidation: builderSummaryResult.ok
      ? builderSummaryResult.validation
      : {
        status: "invalid",
        contract: getBuilderSummaryResponseSchemaExample().contract,
        failure: builderSummaryResult.failure,
      },
    builderRaw: response.summaryContent || "",
  };
}

async function judgeSlide(judgeRuntime, judgeInstructions, config, promptObj, pngPath, context = {}) {
  const familyId = resolveArchetypeFamily(config.archetype);
  const requestEnvelope = createJudgmentRequestEnvelope({
    promptId: promptObj.id,
    prompt: promptObj.text,
    archetype: config.archetype,
    archetypeFamily: familyId,
    expectedVariant: promptObj.expectedVariant || null,
    pngPath,
    builderCarryover: context.builderCarryover || [],
    reviewerCarryover: context.reviewerCarryover || [],
  });

  const variantContext = promptObj.expectedVariant
    ? `\nEXPECTED VARIANT: "${promptObj.expectedVariant}" within the "${familyId}" archetype family.
Score variantMatch=2 if the builder used this specific variant pattern. Score 1 if the family is correct but variant is wrong. Score 0 if the family is also wrong.`
    : `\nNo specific variant expected — score variantMatch based on whether the builder chose the most appropriate variant for this prompt.`;

  const prompt = `${judgeInstructions}

EVALUATE this ${config.archetype} slide.
PNG_PATH: ${pngPath}
ORIGINAL PROMPT: "${promptObj.text}"${variantContext}

${formatRollingContext(context.builderCarryover || [], "Builder carry-over available to the reviewer")}
${formatRollingContext(context.reviewerCarryover || [], "Reviewer carry-over from earlier loops")}
Structured carry-over envelope for this review turn:
${formatCarryoverEnvelope(context.builderCarryover || [], context.reviewerCarryover || [])}

You must inspect the PNG file at the exact path above. Do not infer the review from the prompt alone.
Look specifically at title wording, title/content/source zones, hierarchy, density, whether claims are backed by visible evidence, AND visual execution quality (overlaps, readability, space utilization).
Then score using your 9 dimensions (0-2 each, max 18).
Return a JSON object only. No prose, no markdown.

REQUEST ENVELOPE:
${formatProtocolExample(requestEnvelope)}

RESPONSE CONTRACT:
${formatProtocolExample(getJudgeResponseSchemaExample())}`;

  const response = judgeRuntime
    ? await executeRuntimeRequest(judgeRuntime, {
      type: "prompt",
      prompt,
      timeoutMs: config.judge.timeoutMs || 120000,
      workerTimeoutMs: (config.judge.timeoutMs || 120000) + 30000,
    })
    : await executeAgentRequest(config.judge, {
      type: "prompt",
      prompt,
      timeoutMs: config.judge.timeoutMs || 120000,
      workerTimeoutMs: (config.judge.timeoutMs || 120000) + 30000,
    });

  if (!response.ok) {
    const failure = createFailureDetails(response.error, {
      fallbackCategory: FAILURE_CATEGORIES.toolFailure,
    });
    return {
      ok: false,
      error: failure.message,
      errorCategory: failure.category,
      failure,
      requestEnvelope,
    };
  }

  const text = response.content || "";
  const parsed = parseJudgeResponse(text, { allowLegacyFallback: true });
  if (!parsed.ok) {
    const failure = createFailureDetails(parsed.failure.message, {
      fallbackCategory: parsed.failure.category,
    });
    return {
      ok: false,
      error: failure.message,
      errorCategory: failure.category,
      failure,
      requestEnvelope,
      validation: {
        status: "invalid",
        contract: getJudgeResponseSchemaExample().contract,
        failure: parsed.failure,
      },
      raw: text,
    };
  }

  return {
    ok: true,
    requestEnvelope,
    score: parsed.value.totalScore,
    max: parsed.value.maxScore,
    gaps: parsed.value.gaps.join("; "),
    gapItems: parsed.value.gaps,
    summary: parsed.value.summary,
    archetypeUsed: parsed.value.archetypeUsed,
    archetypeExpected: parsed.value.archetypeExpected,
    dimensionScores: parsed.value.dimensionScores,
    parsedPayload: parsed.value,
    validation: parsed.validation,
    raw: text,
  };
}

async function improveSkills(config, gaps, promptId, loopNumber) {
  if (!config.improver) return { ok: false, error: "no improver configured" };

  const instructionsFile = loadInstructionsFile({
    baseDir: __dirname,
    instructionsFile: config.improver.instructionsFile,
    label: "improver instructions",
  });

  if (!instructionsFile.ok) {
    return {
      ok: false,
      error: instructionsFile.message,
      errorCategory: instructionsFile.category,
    };
  }

  const instructions = instructionsFile.text
    .replace("{SKILLS_DIR}", SKILLS_DIR)
    .replace(/{ARCHETYPES_DIR}/g, ARCHETYPES_DIR);

  const prompt = `${instructions}

ARCHETYPE: ${config.archetype}
LOOP: ${loopNumber}
PROMPT THAT WAS EVALUATED: "${promptId}"

JUDGE GAPS TO FIX:
${gaps}

Read the relevant skill file(s) in ${SKILLS_DIR}, make ONE targeted edit to address the most impactful gap, then report what you changed.`;

  const response = await executeAgentRequest(config.improver, {
    type: "prompt",
    prompt,
    timeoutMs: config.improver.timeoutMs || 180000,
    workerTimeoutMs: (config.improver.timeoutMs || 180000) + 30000,
  });

  return response.ok
    ? { ok: true, response: response.content || "" }
    : (() => {
      const failure = createFailureDetails(response.error, {
        fallbackCategory: FAILURE_CATEGORIES.toolFailure,
      });
      return {
        ok: false,
        error: failure.message,
        errorCategory: failure.category,
        failure,
      };
    })();
}

function createArchetypeBuildFailure(loopState, buildResult) {
  return {
    loop: loopState.loopNumber,
    prompt_id: loopState.prompt.id,
    build_ok: false,
    error: buildResult.error,
    error_category: buildResult.errorCategory,
    error_disposition: buildResult.failure?.disposition || null,
    retry_attempts: buildResult.retry?.attempts || 1,
  };
}

function createArchetypeArtifactFailure(loopState, buildResult, artifactFailure) {
  return {
    loop: loopState.loopNumber,
    prompt_id: loopState.prompt.id,
    build_ok: true,
    judge_ok: false,
    error: artifactFailure?.message || "PNG missing",
    error_category: artifactFailure?.category || FAILURE_CATEGORIES.artifactMissing,
    error_disposition: artifactFailure?.disposition || null,
    builder_contract: buildResult?.builderValidation,
  };
}

function createArchetypeJudgeFailure(loopState, buildResult, judgeResult) {
  return {
    loop: loopState.loopNumber,
    prompt_id: loopState.prompt.id,
    build_ok: true,
    judge_ok: false,
    error: judgeResult.error,
    error_category: judgeResult.errorCategory,
    error_disposition: judgeResult.failure?.disposition || null,
    retry_attempts: judgeResult.retry?.attempts || 1,
    builder_contract: buildResult.builderValidation,
    judge_contract: judgeResult.validation,
    judge_raw: judgeResult.raw,
  };
}

function createArchetypeSuccess(loopState, buildResult, judgeResult) {
  return {
    loop: loopState.loopNumber,
    prompt_id: loopState.prompt.id,
    build_ok: true,
    judge_ok: true,
    score: judgeResult.score,
    max_score: judgeResult.max,
    gaps: judgeResult.gaps,
    judge_summary: judgeResult.summary,
    judge_archetype_used: judgeResult.archetypeUsed,
    judge_archetype_expected: judgeResult.archetypeExpected,
    judge_dimension_scores: judgeResult.dimensionScores,
    builder_summary: buildResult.builderSummary,
    builder_contract: buildResult.builderValidation,
    builder_raw: buildResult.builderRaw,
    judge_contract: judgeResult.validation,
    judge_raw: judgeResult.raw,
    png_path: loopState.artifacts.pngPath,
  };
}

function isArtifactFailureCategory(category) {
  return (
    category === FAILURE_CATEGORIES.artifactMissing
    || category === FAILURE_CATEGORIES.artifactInvalid
  );
}

function getRetryCount(retry) {
  return Math.max(0, (retry?.attempts || 1) - 1);
}

async function persistArchetypeLoopRecord({
  ledger,
  config,
  loopState,
  scoreHistory,
}) {
  const finalResult = loopState.finalResult || {};
  const isBuildFailure = finalResult.build_ok === false;
  const isArtifactFailure = isArtifactFailureCategory(finalResult.error_category);
  const isJudgeFailure = finalResult.build_ok === true && finalResult.judge_ok === false && !isArtifactFailure;
  const improveStepStatus = loopState.stepResults?.[ORCHESTRATOR_STEPS.improve]?.status;
  const skillSnapshotAfter = improveStepStatus && improveStepStatus !== ORCHESTRATOR_STEP_STATUS.skipped
    ? await captureSkillSnapshot({
      baseDir: SKILLS_DIR,
      files: config.builder?.skillFiles || [],
      label: "builder-skills",
    })
    : null;

  await ledger.writeLoopRecord({
    loopNumber: loopState.loopNumber,
    prompt: {
      id: loopState.prompt.id,
      text: loopState.prompt.text,
      category: config.archetype,
      archetype: config.archetype,
    },
    status: isBuildFailure
      ? "build_failed"
      : isArtifactFailure
        ? "artifact_failed"
        : isJudgeFailure
          ? "judge_failed"
          : "completed",
    buildStatus: isBuildFailure ? "failed" : "succeeded",
    judgeStatus: isBuildFailure || isArtifactFailure ? "not_run" : isJudgeFailure ? "failed" : "succeeded",
    score: finalResult.score ?? null,
    maxScore: finalResult.max_score ?? null,
    errorCategory: finalResult.error_category ?? null,
    errorMessage: finalResult.error ?? null,
    timings: {
      startedAt: loopState.startedAt,
      finishedAt: loopState.completedAt,
      phases: loopState.metadata.stepTimings,
    },
    retries: {
      build: getRetryCount(loopState.buildResult?.retry),
      judge: getRetryCount(loopState.judgeResult?.retry),
      improvement: 0,
    },
    carryover: loopState.carryoverSnapshot,
    diagnostics: {
      cleanupSessionsClosed: loopState.metadata.cleanupCount || 0,
      cleanupFailures: loopState.metadata.cleanupFailures || [],
      stepHistory: loopState.stepHistory,
      scoreHistory: [...scoreHistory],
      recovery: loopState.recovery,
    },
    builder: {
      request: loopState.buildResult?.requestEnvelope || null,
      completion: loopState.buildResult?.completion || null,
      summary: loopState.buildResult?.builderSummary || null,
      raw: loopState.buildResult?.builderRaw || null,
      errorMessage: loopState.buildResult?.error || null,
      errorCategory: loopState.buildResult?.errorCategory || null,
      validation: loopState.buildResult?.builderValidation || null,
      retry: loopState.buildResult?.retry || null,
    },
    judge: {
      request: loopState.judgeResult?.requestEnvelope || null,
      parsed: loopState.judgeResult?.ok ? {
        totalScore: loopState.judgeResult.score,
        maxScore: loopState.judgeResult.max,
        gaps: loopState.judgeResult.gaps,
        summary: loopState.judgeResult.summary,
        archetypeUsed: loopState.judgeResult.archetypeUsed,
        archetypeExpected: loopState.judgeResult.archetypeExpected,
        dimensionScores: loopState.judgeResult.dimensionScores,
      } : null,
      summary: loopState.judgeResult?.summary || null,
      raw: loopState.judgeResult?.raw || null,
      errorMessage: loopState.judgeResult?.error || null,
      errorCategory: loopState.judgeResult?.errorCategory || null,
      validation: loopState.judgeResult?.validation || null,
      retry: loopState.judgeResult?.retry || null,
    },
    improvement: loopState.improveResult?.ok
      ? { attempted: true, response: loopState.improveResult.response || "" }
      : loopState.improveResult
        ? { attempted: true, errorMessage: loopState.improveResult.error || "" }
        : { attempted: false },
    artifacts: [
      { role: "png", kind: "image", path: loopState.artifacts.pngPath },
      { role: "pptx", kind: "presentation", path: loopState.artifacts.pptxPath },
    ],
    skills: {
      beforeLoop: loopState.metadata.persistence?.skillSnapshotBefore || null,
      afterLoop: skillSnapshotAfter,
    },
    metadata: {
      configName: config.name,
      runName: config.runName,
      mode: config.mode,
      goal: config.goal,
      builderModel: config.builder?.model,
      judgeModel: config.judge?.model,
    },
  });
}

async function runConfig(configPath) {
  const rawConfig = JSON.parse(readFileSync(configPath, "utf-8"));
  const mode = resolveEvalMode({
    configuredMode: rawConfig.mode,
    hasImprover: Boolean(rawConfig.improver),
  });
  const improverDisabledInBaseline = mode === EVAL_MODES.baseline && Boolean(rawConfig.improver);
  const builderTransport = normalizeTransport(rawConfig.builder?.transport || "cli");
  const config = {
    ...rawConfig,
    mode,
    runName: getModeTaggedName(rawConfig.name, mode),
    improver: mode === EVAL_MODES.baseline ? null : rawConfig.improver ?? null,
  };
  const resultsDir = getModeScopedDirectory(resolveEvalAssetPathFromRelative(config.resultsDir), config.mode);
  const outputDir = getModeScopedDirectory(resolveEvalAssetPathFromRelative(config.outputDir), config.mode);
  mkdirSync(resultsDir, { recursive: true });
  mkdirSync(outputDir, { recursive: true });

  const ledger = await createEvalLedger({
    ledgerRoot: join(resultsDir, "ledger"),
    runName: config.runName,
    runner: "run-archetype-eval",
    metadata: {
      configPath,
      config: {
        name: config.name,
        runName: config.runName,
        mode: config.mode,
        archetype: config.archetype,
        loops: config.loops,
        goal: config.goal,
        builder: {
          model: config.builder?.model,
          transport: builderTransport,
          skillFiles: config.builder?.skillFiles || [],
        },
        judge: {
          model: config.judge?.model,
        },
        improver: config.improver
          ? {
            model: config.improver.model,
          }
          : null,
      },
    },
  });

  console.log(`\n${"═".repeat(60)}`);
  console.log(`🎯 ${config.runName}: ${config.description}`);
  console.log(`   Mode: ${config.mode}`);
  console.log(`   Goal: ${config.goal}`);
  console.log(`   Loops: ${config.loops} | Prompts: ${config.prompts.length}`);
  const builderLabel = config.builder.reasoningEffort
    ? `${config.builder.model} (${config.builder.reasoningEffort})`
    : config.builder.model;
  const judgeLabel = config.judge.reasoningEffort
    ? `${config.judge.model} (${config.judge.reasoningEffort})`
    : config.judge.model;
  const transportLabel = builderTransport === "mcp" ? " via MCP" : "";
  const isolationLabel = shouldUseIsolatedProcess(config.builder) ? " [isolated]" : "";
  const judgeIsolationLabel = shouldUseIsolatedProcess(config.judge) ? " [isolated]" : "";
  console.log(`   Builder: ${builderLabel}${transportLabel}${isolationLabel} | Judge: ${judgeLabel}${judgeIsolationLabel}`);
  if (config.improver) {
    const improverLabel = config.improver.reasoningEffort
      ? `${config.improver.model} (${config.improver.reasoningEffort})`
      : config.improver.model;
    console.log(`   Improver: ${improverLabel}${shouldUseIsolatedProcess(config.improver) ? " [isolated]" : ""}`);
  } else if (improverDisabledInBaseline) {
    console.log("   Improver: ignored in baseline mode");
  }
  console.log(`${"═".repeat(60)}\n`);

  const serviceStatus = ensureService();
  if (!serviceStatus.ok) {
    throw new Error(`Failed to start pptcli service: ${serviceStatus.error}`);
  }

  const judgeInstructionsFile = loadInstructionsFile({
    baseDir: __dirname,
    instructionsFile: config.judge.instructionsFile,
    label: "judge instructions",
  });
  if (!judgeInstructionsFile.ok) {
    throw new Error(judgeInstructionsFile.message);
  }
  const judgeInstructions = judgeInstructionsFile.text;

  let judgeRuntime = null;
  let builderRuntime = null;

  if (shouldUseIsolatedProcess(config.judge)) {
    console.log(`  📐 Judge will run in isolated process mode\n`);
  } else {
    console.log(`  📐 Starting judge session (${judgeLabel})...`);
    judgeRuntime = await createRuntime(config.judge, { executionMode: "reuse-session" });
    await executeRuntimeRequest(judgeRuntime, {
      type: "prompt",
      prompt: `${judgeInstructions}\n\nYou will evaluate ${config.archetype} slides. Always return JSON objects only. Acknowledge with JSON {"acknowledged": true}.`,
      timeoutMs: 60000,
    });
    console.log(`  ✅ Judge ready\n`);
  }

  if (shouldReuseSessionContext(config.builder)) {
    console.log(`  🧠 Reusing builder session context across loops\n`);
    builderRuntime = await createRuntime(config.builder);
  }

  try {
    const orchestrationContext = createOrchestrationContext({
      builderRuntime,
      judgeRuntime,
      builderCarryover: [],
      reviewerCarryover: [],
      scoreHistory: [],
      metadata: {
        configName: config.name,
        loops: config.loops,
      },
    });

    const items = Array.from({ length: config.loops }, (_, index) => ({
      loopNumber: index + 1,
      prompt: config.prompts[index % config.prompts.length],
    }));

    const { results } = await runLoopOrchestrator({
      items,
      context: orchestrationContext,
      createLoopState: ({ item, context }) => {
        const pngPath = join(outputDir, `loop${item.loopNumber}-${item.prompt.id}.png`);

        return createLoopState({
          loopNumber: item.loopNumber,
          prompt: item.prompt,
          pngPath,
          pptxPath: pngPath.replace(".png", ".pptx"),
          sequence: createStepSequence({
            includeJudge: true,
            includeImprove: Boolean(config.improver),
          }),
            carryoverSnapshot: context.carryover,
            metadata: {
              configName: config.name,
              runName: config.runName,
              mode: config.mode,
              archetype: config.archetype,
            },
          });
      },
      shouldSkipStep: (step, loopState) => (
        step === ORCHESTRATOR_STEPS.improve
        && (!config.improver || !loopState.finalResult?.gaps)
      ),
      handlers: {
        [ORCHESTRATOR_STEPS.cleanup]: async (loopState) => {
          const cleanup = cleanupPptSessions();
          loopState.metadata.cleanupCount = cleanup.count;
          loopState.metadata.cleanupFailures = cleanup.failedSessionIds || [];
          return {
            status: cleanup.ok === false ? ORCHESTRATOR_STEP_STATUS.failure : ORCHESTRATOR_STEP_STATUS.success,
            failure: cleanup.ok === false
              ? createFailureDetails(cleanup.error || "Cleanup failed", {
                fallbackCategory: cleanup.errorCategory || FAILURE_CATEGORIES.cleanupFailure,
              })
              : null,
          };
        },
        [ORCHESTRATOR_STEPS.build]: async (loopState, context) => {
          console.log(`  Loop ${loopState.loopNumber}/${config.loops} — ${loopState.prompt.id}`);
          console.log(`    🔨 Building...`);

          const buildResult = await executeWithRetry(
            () => buildSlide(
              config,
              loopState.prompt,
              loopState.artifacts.pngPath,
              context.runtimes.builder,
              {
                builderCarryover: context.carryover.builder,
                reviewerCarryover: context.carryover.reviewer,
              }
            ),
            {
              ...DEFAULT_RETRY_POLICY,
              classifyFailure: (result) => result?.failure || createFailureDetails(result?.error, {
                fallbackCategory: result?.errorCategory,
              }),
              onRetry: async ({ failure, attempt, maxAttempts, delayMs }) => {
                console.log(`    ↻ Retrying build (${attempt + 1}/${maxAttempts}) in ${delayMs}ms: ${failure.message}`);
              },
            }
          );
          loopState.buildResult = buildResult;

          if (!buildResult.ok) {
            console.log(`    ❌ Build failed: ${buildResult.error.slice(0, 80)}`);
            const finalResult = createArchetypeBuildFailure(loopState, buildResult);
            loopState.finalResult = finalResult;

            return {
              status: ORCHESTRATOR_STEP_STATUS.failure,
              terminate: true,
              finalResult,
              terminationReason: "build-failure",
              failure: buildResult.failure,
              retry: buildResult.retry,
            };
          }

          console.log(`    ✅ Built`);
          if (buildResult.builderValidation?.status !== "strict") {
            console.log(`    ⚠️  Builder summary contract ${buildResult.builderValidation?.status || "invalid"}${buildResult.builderValidation?.failure?.category ? ` (${buildResult.builderValidation.failure.category})` : ""}`);
          }

          if (buildResult.builderSummary) {
            pushRollingItem(
              context.carryover.builder,
              createBuilderCarryoverEntry({
                loopNumber: loopState.loopNumber,
                promptId: loopState.prompt.id,
                prompt: loopState.prompt.text,
                summary: buildResult.builderSummary,
                validation: buildResult.builderValidation,
              })
            );
          }

          return { status: ORCHESTRATOR_STEP_STATUS.success };
        },
        [ORCHESTRATOR_STEPS.verifyArtifact]: async (loopState) => {
          const artifactStatus = await verifyBuildArtifacts({
            pngPath: loopState.artifacts.pngPath,
            pptxPath: loopState.artifacts.pptxPath,
            requirePptx: true,
            timeoutMs: 3500,
          });

          if (!artifactStatus.ok) {
            console.log(`    ⚠️  Artifact check failed: ${artifactStatus.message}`);
            const failure = createFailureDetails(artifactStatus.message, {
              fallbackCategory: artifactStatus.category,
            });
            const finalResult = createArchetypeArtifactFailure(loopState, loopState.buildResult, failure);
            loopState.finalResult = finalResult;

            return {
              status: ORCHESTRATOR_STEP_STATUS.failure,
              terminate: true,
              finalResult,
              terminationReason: "artifact-missing",
              failure,
            };
          }

          return { status: ORCHESTRATOR_STEP_STATUS.success };
        },
        [ORCHESTRATOR_STEPS.judge]: async (loopState, context) => {
          console.log(`    ⚖️  Judging...`);

          // Helper: attempt judge with optional session recovery
          const attemptJudge = async () => {
            return judgeSlide(
              context.runtimes.judge,
              judgeInstructions,
              config,
              loopState.prompt,
              loopState.artifacts.pngPath,
              {
                builderCarryover: context.carryover.builder,
                reviewerCarryover: context.carryover.reviewer,
              }
            );
          };

          let judgeResult = await executeWithRetry(
            attemptJudge,
            {
              ...DEFAULT_RETRY_POLICY,
              classifyFailure: (result) => result?.failure || createFailureDetails(result?.error, {
                fallbackCategory: result?.errorCategory,
              }),
              onRetry: async ({ failure, attempt, maxAttempts, delayMs }) => {
                console.log(`    ↻ Retrying judge (${attempt + 1}/${maxAttempts}) in ${delayMs}ms: ${failure.message}`);
              },
            }
          );

          // If judge failed due to dead session, try to recreate it once
          if (!judgeResult.ok && judgeResult.error && judgeResult.error.includes("Session not found") && !shouldUseIsolatedProcess(config.judge)) {
            console.log(`    🔄 Judge session died — recreating...`);
            try {
              await destroyRuntime(context.runtimes.judge);
            } catch {}
            try {
              context.runtimes.judge = await createRuntime(config.judge, { executionMode: "reuse-session" });
              await executeRuntimeRequest(context.runtimes.judge, {
                type: "prompt",
                prompt: `${judgeInstructions}\n\nYou will evaluate ${config.archetype} slides. Always return JSON objects only. Acknowledge with JSON {"acknowledged": true}.`,
                timeoutMs: 60000,
              });
              console.log(`    ✅ Judge session recovered`);

              judgeResult = await executeWithRetry(
                attemptJudge,
                {
                  ...DEFAULT_RETRY_POLICY,
                  classifyFailure: (result) => result?.failure || createFailureDetails(result?.error, {
                    fallbackCategory: result?.errorCategory,
                  }),
                  onRetry: async ({ failure, attempt, maxAttempts, delayMs }) => {
                    console.log(`    ↻ Retrying judge after recovery (${attempt + 1}/${maxAttempts}) in ${delayMs}ms: ${failure.message}`);
                  },
                }
              );
            } catch (recoveryError) {
              console.log(`    ❌ Judge recovery failed: ${recoveryError.message?.slice(0, 80) || recoveryError}`);
            }
          }

          loopState.judgeResult = judgeResult;

          if (!judgeResult.ok) {
            console.log(`    ❌ Judge failed: ${judgeResult.error.slice(0, 80)}`);
            const finalResult = createArchetypeJudgeFailure(loopState, loopState.buildResult, judgeResult);
            loopState.finalResult = finalResult;

            return {
              status: ORCHESTRATOR_STEP_STATUS.failure,
              terminate: false,
              finalResult,
              terminationReason: "judge-failure",
              failure: judgeResult.failure,
              retry: judgeResult.retry,
            };
          }

          context.scoreHistory.push(judgeResult.score);
          const averageScore = (context.scoreHistory.reduce((sum, value) => sum + value, 0) / context.scoreHistory.length).toFixed(1);
          const trend = context.scoreHistory.length > 1
            ? (judgeResult.score > context.scoreHistory[context.scoreHistory.length - 2] ? "📈" : judgeResult.score < context.scoreHistory[context.scoreHistory.length - 2] ? "📉" : "➡️")
            : "";

          console.log(`    📊 Score: ${judgeResult.score}/${judgeResult.max} | Avg: ${averageScore} ${trend}`);
          if (judgeResult.validation?.status !== "strict") {
            console.log(`    ⚠️  Judge contract ${judgeResult.validation?.status || "invalid"}${judgeResult.validation?.failure?.category ? ` (${judgeResult.validation.failure.category})` : ""}`);
          }

          pushRollingItem(
            context.carryover.reviewer,
            createReviewerCarryoverEntry({
              loopNumber: loopState.loopNumber,
              promptId: loopState.prompt.id,
              prompt: loopState.prompt.text,
              judgment: judgeResult.parsedPayload,
              validation: judgeResult.validation,
            })
          );

          const finalResult = createArchetypeSuccess(loopState, loopState.buildResult, judgeResult);
          loopState.finalResult = finalResult;

          return {
            status: ORCHESTRATOR_STEP_STATUS.success,
            finalResult,
          };
        },
        [ORCHESTRATOR_STEPS.improve]: async (loopState) => {
          console.log(`    🔧 Improving skills...`);
          const improveResult = await improveSkills(config, loopState.finalResult.gaps, loopState.prompt.id, loopState.loopNumber);
          loopState.improveResult = improveResult;

          if (improveResult.ok) {
            console.log(`    ✅ Skills updated`);
            loopState.finalResult = {
              ...loopState.finalResult,
              improvement: improveResult.response?.slice(0, 200),
            };

            return {
              status: ORCHESTRATOR_STEP_STATUS.success,
              finalResult: loopState.finalResult,
            };
          }

          console.log(`    ⚠️  Improve skipped: ${improveResult.error?.slice(0, 60)}`);
          return {
            status: ORCHESTRATOR_STEP_STATUS.failure,
            improveError: improveResult.error,
            finalResult: loopState.finalResult,
          };
        },
      },
      onLoopStart: async (loopState) => {
        loopState.metadata.persistence = {
          skillSnapshotBefore: await captureSkillSnapshot({
            baseDir: SKILLS_DIR,
            files: config.builder?.skillFiles || [],
            label: "builder-skills",
          }),
        };
      },
      onStepStart: (step, loopState) => {
        loopState.metadata.stepTimings = loopState.metadata.stepTimings || {};
        loopState.metadata.stepTimings[step] = {
          startedAt: new Date().toISOString(),
        };
      },
      onStepComplete: (step, loopState) => {
        const stepTiming = loopState.metadata.stepTimings?.[step];
        if (stepTiming) {
          const finishedAt = new Date().toISOString();
          stepTiming.finishedAt = finishedAt;
          stepTiming.durationMs = Date.parse(finishedAt) - Date.parse(stepTiming.startedAt);
        }
      },
      onLoopComplete: async (loopState, context) => {
        await persistArchetypeLoopRecord({
          ledger,
          config,
          loopState,
          scoreHistory: context.scoreHistory,
        });
        console.log();
        return loopState.finalResult;
      },
    });

    const reportPath = join(ledger.runDir, "run-report.json");
    const { report } = await generateRunReportFromPersistence({
      manifestPath: ledger.manifestPath,
      transactionsPath: ledger.transactionsPath,
      outputPath: reportPath,
    });

    console.log(`\n${formatRunReportForConsole(report)}`);
    if (orchestrationContext.scoreHistory.length > 1) {
      console.log(`  Score history: ${orchestrationContext.scoreHistory.join(" → ")}`);
    }

    const timestamp = new Date().toISOString().slice(0, 19).replace(/[T:]/g, "-");
    const outFile = join(resultsDir, `${config.runName}-${timestamp}.json`);
    writeFileSync(outFile, JSON.stringify({
      config: {
        name: config.name,
        runName: config.runName,
        mode: config.mode,
        archetype: config.archetype,
        loops: config.loops,
        goal: config.goal,
      },
      summary: {
        ...report.summary,
        mode: config.mode,
        scored: report.summary.scoredLoops,
        avg: report.summary.averageScore,
        score_history: orchestrationContext.scoreHistory,
      },
      report,
      results,
    }, null, 2));
    await ledger.finalize({
      summary: {
        ...report.summary,
        mode: config.mode,
        scored: report.summary.scoredLoops,
        avg: report.summary.averageScore,
        scoreHistory: orchestrationContext.scoreHistory,
      },
      finalArtifacts: [
        { role: "results-json", kind: "record", path: outFile },
        { role: "run-report", kind: "record", path: reportPath },
      ],
    });
    console.log(`💾 ${outFile}\n`);

    return {
      name: config.runName,
      avg: report.summary.averageScore || 0,
      scored: report.summary.scoredLoops,
      total: config.loops,
      history: orchestrationContext.scoreHistory,
    };
  } finally {
    await destroyRuntime(judgeRuntime);
    await destroyRuntime(builderRuntime);
  }
}

async function main() {
  const arg = process.argv[2];
  if (!arg) {
    console.log("Usage:");
    console.log("  node run-archetype-eval.mjs configs/<archetype>.json");
    console.log("  node run-archetype-eval.mjs configs/all");
    process.exit(1);
  }

  if (arg.endsWith("/all") || arg.endsWith("\\all")) {
    const configDir = join(__dirname, "configs");
    const files = readdirSync(configDir).filter((file) => file.endsWith(".json") && !file.startsWith("_")).sort();
    console.log(`\n🚀 Running ${files.length} archetype evaluations\n`);
    const summaries = [];
    for (const file of files) {
      summaries.push(await runConfig(join(configDir, file)));
    }

    console.log(`${"═".repeat(60)}`);
    console.log("🏁 ALL ARCHETYPES");
    for (const summary of summaries) {
      console.log(`   ${summary.name.padEnd(20)} avg: ${String(summary.avg).padStart(4)} | ${summary.history.join("→") || "no scores"}`);
    }
    return;
  }

  await runConfig(arg);
}

main().catch((err) => {
  console.error("Fatal:", err);
  process.exit(1);
});
