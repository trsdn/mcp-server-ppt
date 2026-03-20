/**
 * Automated Slide Design Evaluation Loop
 *
 * Architecture:
 *   Builder Agent (Sonnet) + PptMcp MCP Server + Design Skills → builds slide → exports PNG
 *   Judge Agent (GPT) + Structure Instructions → evaluates PNG → scores + gaps
 *   Harness → records results, applies gaps to skills, repeats
 *
 * Usage:
 *   node run-eval.mjs                    # Run all 110 prompts
 *   node run-eval.mjs --start 0 --count 5  # Run first 5
 *   node run-eval.mjs --category dashboard  # Run only dashboard prompts
 */

import { readFileSync, writeFileSync, mkdirSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import {
  createRuntime,
  destroyRuntime,
  executeAgentRequest,
  executeWithRetry,
  loadInstructionsFile,
  executeRuntimeRequest,
  verifyBuildArtifacts,
} from "./lib/runtime/index.mjs";
import { CLI_PATH, EVAL_OUTPUT_ROOT, EVAL_RESULTS_ROOT, SKILLS_DIR } from "./lib/runtime/environment.mjs";
import {
  ensureEvalMode,
  getModeScopedDirectory,
  getModeTaggedName,
} from "./lib/mode.mjs";
import {
  FAILURE_CATEGORIES,
  createFailureDetails,
  createEvaluationRequestEnvelope,
  createJudgmentRequestEnvelope,
  formatProtocolExample,
  getJudgeResponseSchemaExample,
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
  runLoopOrchestrator,
} from "./lib/orchestrator/index.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROMPTS_FILE = join(__dirname, "prompts", "test-prompts.json");
const OUTPUT_DIR = EVAL_OUTPUT_ROOT;
const RESULTS_DIR = EVAL_RESULTS_ROOT;
const SKILL_FILES = [
  "slide-design-principles.md",
  "slide-design-review.md",
  "generation-pipeline.md",
];
const DEFAULT_RETRY_POLICY = Object.freeze({
  maxAttempts: 2,
  baseDelayMs: 1500,
});

function loadPrompts() {
  return JSON.parse(readFileSync(PROMPTS_FILE, "utf-8"));
}

function parseArgs() {
  const args = process.argv.slice(2);
  const opts = {
    mode: "baseline",
    start: 0,
    count: Infinity,
    category: null,
    builderModel: "claude-sonnet-4.5",
    builderReasoningEffort: null,
    builderIsolatedProcess: false,
    builderReuseSessionContext: false,
    judgeModel: "gpt-5",
    judgeReasoningEffort: null,
    judgeIsolatedProcess: false,
    judgeReuseSessionContext: false,
  };

  for (let i = 0; i < args.length; i++) {
    if (args[i] === "--start") opts.start = parseInt(args[++i]);
    if (args[i] === "--count") opts.count = parseInt(args[++i]);
    if (args[i] === "--category") opts.category = args[++i];
    if (args[i] === "--mode") opts.mode = ensureEvalMode(args[++i], "--mode");
    if (args[i] === "--builder-model") opts.builderModel = args[++i];
    if (args[i] === "--builder-reasoning-effort") opts.builderReasoningEffort = args[++i];
    if (args[i] === "--builder-isolated-process") opts.builderIsolatedProcess = true;
    if (args[i] === "--builder-reuse-session-context") opts.builderReuseSessionContext = true;
    if (args[i] === "--judge-model") opts.judgeModel = args[++i];
    if (args[i] === "--judge-reasoning-effort") opts.judgeReasoningEffort = args[++i];
    if (args[i] === "--judge-isolated-process") opts.judgeIsolatedProcess = true;
    if (args[i] === "--judge-reuse-session-context") opts.judgeReuseSessionContext = true;
  }

  return opts;
}

function getBuilderAgent(opts) {
  return {
    model: opts.builderModel,
    reasoningEffort: opts.builderReasoningEffort,
    isolatedProcess: opts.builderIsolatedProcess,
    reuseSessionContext: opts.builderReuseSessionContext,
  };
}

function getJudgeAgent(opts) {
  return {
    model: opts.judgeModel,
    reasoningEffort: opts.judgeReasoningEffort,
    isolatedProcess: opts.judgeIsolatedProcess,
    reuseSessionContext: opts.judgeReuseSessionContext,
  };
}

async function runBuilder(prompt, outputPath, opts, builderRuntime = null) {
  const builderInstructionsFile = loadInstructionsFile({
    baseDir: __dirname,
    instructionsFile: join("agents", "builder-instructions.md"),
    expectedTransport: "cli",
    label: "builder instructions",
  });

  if (!builderInstructionsFile.ok) {
    const failure = createFailureDetails(builderInstructionsFile.message, {
      fallbackCategory: builderInstructionsFile.category,
    });
    return {
      success: false,
      error: failure.message,
      errorCategory: failure.category,
      failure,
      requestEnvelope: null,
      pptxPath: outputPath.replace(".png", ".pptx"),
    };
  }

  const builderInstructions = builderInstructionsFile.text.replace("{CLI_PATH}", CLI_PATH);
  const pptxPath = outputPath.replace(".png", ".pptx");
  const requestEnvelope = createEvaluationRequestEnvelope({
    promptId: prompt.id,
    prompt: prompt.prompt,
    archetype: prompt.category,
    transport: "cli",
    pngPath: outputPath,
    pptxPath,
  });
  const skillPaths = SKILL_FILES.map((file) => join(SKILLS_DIR, file)).join("\n- ");
  const buildPrompt = `${builderInstructions}

You build PowerPoint slides using the pptcli CLI.

CLI: ${CLI_PATH}
RULES: --color not --font-color. --alignment not --horizontal-alignment. No \\n in --text. Service running. Close existing sessions first.

FIRST: Read these design skill files for guidance:
- ${skillPaths}

THEN: Build ONE slide for this request:
"${prompt.prompt}" (category: ${prompt.category})

OUTPUT: ${pptxPath} (PPTX) and ${outputPath} (PNG export)

Steps: read skills → pick archetype → create session → build slide → export PNG → close --save.
`;

  const response = builderRuntime
    ? await executeRuntimeRequest(builderRuntime, {
      type: "prompt",
      prompt: buildPrompt,
      timeoutMs: 300000,
    })
    : await executeAgentRequest(getBuilderAgent(opts), {
      type: "prompt",
      prompt: buildPrompt,
      timeoutMs: 300000,
    });

  return response.ok
    ? {
      success: true,
      response: response.content || "completed",
      requestEnvelope,
      pptxPath,
    }
    : (() => {
      const failure = createFailureDetails(response.error, {
        fallbackCategory: FAILURE_CATEGORIES.toolFailure,
      });
      return {
        success: false,
        error: failure.message,
        errorCategory: failure.category,
        failure,
        requestEnvelope,
        pptxPath,
      };
    })();
}

async function runJudge(prompt, pngPath, opts, judgeRuntime = null) {
  const judgeInstructionsFile = loadInstructionsFile({
    baseDir: __dirname,
    instructionsFile: join("agents", "judge-instructions.md"),
    label: "judge instructions",
  });

  if (!judgeInstructionsFile.ok) {
    const failure = createFailureDetails(judgeInstructionsFile.message, {
      fallbackCategory: judgeInstructionsFile.category,
    });
    return {
      success: false,
      error: failure.message,
      errorCategory: failure.category,
      failure,
      requestEnvelope: null,
    };
  }

  const judgeInstructions = judgeInstructionsFile.text;
  const requestEnvelope = createJudgmentRequestEnvelope({
    promptId: prompt.id,
    prompt: prompt.prompt,
    archetype: prompt.category,
    pngPath,
  });
  const judgePrompt = `
${judgeInstructions}

Evaluate this slide image for structural correctness.

ORIGINAL PROMPT: "${prompt.prompt}"
CATEGORY: ${prompt.category}
IMAGE PATH: ${pngPath}

View the image file and score it using the 7 dimensions from your instructions.
Return JSON matching this contract:
${formatProtocolExample(getJudgeResponseSchemaExample())}

Request envelope:
${formatProtocolExample(requestEnvelope)}
`;

  const response = judgeRuntime
    ? await executeRuntimeRequest(judgeRuntime, {
      type: "prompt",
      prompt: judgePrompt,
      timeoutMs: 120000,
    })
    : await executeAgentRequest(getJudgeAgent(opts), {
      type: "prompt",
      prompt: judgePrompt,
      timeoutMs: 120000,
    });

  if (!response.ok) {
    const failure = createFailureDetails(response.error, {
      fallbackCategory: FAILURE_CATEGORIES.toolFailure,
    });
    return {
      success: false,
      error: failure.message,
      errorCategory: failure.category,
      failure,
      requestEnvelope,
    };
  }

  const parsed = parseJudgeResponse(response.content || "", { allowLegacyFallback: true });
  if (!parsed.ok) {
    const failure = createFailureDetails(parsed.failure.message, {
      fallbackCategory: parsed.failure.category,
    });
    return {
      success: false,
      error: failure.message,
      errorCategory: failure.category,
      failure,
      requestEnvelope,
      validation: {
        status: "invalid",
        contract: getJudgeResponseSchemaExample().contract,
        failure: parsed.failure,
      },
      raw: response.content || "",
    };
  }

  return {
    success: true,
    requestEnvelope,
    parsed: parsed.value,
    validation: parsed.validation,
    raw: response.content || "",
  };
}

function createBuildFailureResult(prompt, buildResult) {
  return {
    prompt_id: prompt.id,
    category: prompt.category,
    build_success: false,
    error: buildResult.error,
    error_category: buildResult.errorCategory,
    error_disposition: buildResult.failure?.disposition || null,
    retry_attempts: buildResult.retry?.attempts || 1,
  };
}

function createArtifactMissingResult(prompt, artifactFailure) {
  return {
    prompt_id: prompt.id,
    category: prompt.category,
    build_success: true,
    judge_success: false,
    error: artifactFailure?.message || "PNG not exported",
    error_category: artifactFailure?.category || FAILURE_CATEGORIES.artifactMissing,
    error_disposition: artifactFailure?.disposition || null,
  };
}

function createJudgeFailureResult(prompt, judgeResult) {
  return {
    prompt_id: prompt.id,
    category: prompt.category,
    build_success: true,
    judge_success: false,
    error: judgeResult.error,
    error_category: judgeResult.errorCategory,
    error_disposition: judgeResult.failure?.disposition || null,
    retry_attempts: judgeResult.retry?.attempts || 1,
    judge_contract: judgeResult.validation,
    judge_raw: judgeResult.raw,
  };
}

function createSuccessResult(prompt, judgeResult, pngPath) {
  const gaps = judgeResult.parsed.gaps.join("; ");

  return {
    prompt_id: prompt.id,
    category: prompt.category,
    build_success: true,
    judge_success: true,
    score: judgeResult.parsed.totalScore,
    max_score: judgeResult.parsed.maxScore,
    gaps,
    judge_contract: judgeResult.validation,
    judge_dimension_scores: judgeResult.parsed.dimensionScores,
    judge_summary: judgeResult.parsed.summary,
    judge_raw: judgeResult.raw,
    png_path: pngPath,
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

async function main() {
  const opts = parseArgs();
  const runName = getModeTaggedName("auto-eval", opts.mode);
  const modeOutputDir = getModeScopedDirectory(OUTPUT_DIR, opts.mode);
  const modeResultsDir = getModeScopedDirectory(RESULTS_DIR, opts.mode);
  const prompts = loadPrompts();
  const serviceStatus = ensureService();

  if (!serviceStatus.ok) {
    throw new Error(`Failed to start pptcli service: ${serviceStatus.error}`);
  }

  let selected = prompts;
  if (opts.category) {
    selected = selected.filter((prompt) => prompt.category === opts.category);
  }
  selected = selected.slice(opts.start, opts.start + opts.count);

  console.log(`\n🔄 Slide Design Eval Loop`);
  console.log(`   Mode:    ${opts.mode}`);
  console.log(`   Prompts: ${selected.length} (of ${prompts.length})`);
  console.log(`   Builder: ${opts.builderModel}${opts.builderReasoningEffort ? ` (${opts.builderReasoningEffort})` : ""}${opts.builderIsolatedProcess ? " [isolated]" : opts.builderReuseSessionContext ? " [reuse]" : ""}`);
  console.log(`   Judge:   ${opts.judgeModel}${opts.judgeReasoningEffort ? ` (${opts.judgeReasoningEffort})` : ""}${opts.judgeIsolatedProcess ? " [isolated]" : opts.judgeReuseSessionContext ? " [reuse]" : ""}`);
  console.log(`   Output:  ${modeOutputDir}\n`);

  mkdirSync(modeOutputDir, { recursive: true });
  mkdirSync(modeResultsDir, { recursive: true });

  const ledger = await createEvalLedger({
    ledgerRoot: join(modeResultsDir, "ledger"),
    runName,
    runner: "run-eval",
    metadata: {
      mode: opts.mode,
      promptsFile: PROMPTS_FILE,
      selectedPromptIds: selected.map((prompt) => prompt.id),
      skillFiles: SKILL_FILES,
      options: opts,
    },
  });

  let builderRuntime = null;
  let judgeRuntime = null;

  if (opts.builderReuseSessionContext && !opts.builderIsolatedProcess) {
    builderRuntime = await createRuntime(getBuilderAgent(opts));
  }
  if (opts.judgeReuseSessionContext && !opts.judgeIsolatedProcess) {
    judgeRuntime = await createRuntime(getJudgeAgent(opts));
  }

  try {
    const orchestrationContext = createOrchestrationContext({
      builderRuntime,
      judgeRuntime,
      metadata: { opts, totalPrompts: selected.length },
    });

    const { results } = await runLoopOrchestrator({
      items: selected,
      context: orchestrationContext,
      createLoopState: ({ item, index, context }) => {
        const pngPath = join(modeOutputDir, `auto-${item.id}.png`);

        return createLoopState({
          loopNumber: index + 1,
          prompt: item,
          pngPath,
          pptxPath: pngPath.replace(".png", ".pptx"),
          sequence: createStepSequence({ includeJudge: true }),
          carryoverSnapshot: context.carryover,
          metadata: {
            totalPrompts: selected.length,
          },
        });
      },
      handlers: {
        [ORCHESTRATOR_STEPS.cleanup]: async (loopState) => {
          const cleanup = cleanupPptSessions();
          loopState.metadata.cleanupCount = cleanup.count;
          loopState.metadata.cleanupFailures = cleanup.failedSessionIds || [];
          for (const sessionId of cleanup.closedSessionIds) {
            console.log(`  🧹 Closed stale session ${sessionId.slice(0, 8)}...`);
          }
          for (const failedSession of cleanup.failedSessionIds || []) {
            console.log(`  ⚠️ Cleanup failed for ${failedSession.sessionId.slice(0, 8)}...: ${failedSession.error}`);
          }

          return {
            status: cleanup.ok === false ? ORCHESTRATOR_STEP_STATUS.failure : ORCHESTRATOR_STEP_STATUS.success,
            cleanupCount: cleanup.count,
            failure: cleanup.ok === false
              ? createFailureDetails(cleanup.error || "Cleanup failed", {
                fallbackCategory: cleanup.errorCategory || FAILURE_CATEGORIES.cleanupFailure,
              })
              : null,
          };
        },
        [ORCHESTRATOR_STEPS.build]: async (loopState, context) => {
          console.log(`  🔨 Building with ${opts.builderModel}${opts.builderReasoningEffort ? ` (${opts.builderReasoningEffort})` : ""}...`);
          const buildResult = await executeWithRetry(
            () => runBuilder(loopState.prompt, loopState.artifacts.pngPath, opts, context.runtimes.builder),
            {
              ...DEFAULT_RETRY_POLICY,
              isSuccess: (result) => Boolean(result?.success),
              classifyFailure: (result) => result?.failure || createFailureDetails(result?.error, {
                fallbackCategory: result?.errorCategory,
              }),
              onRetry: async ({ failure, attempt, maxAttempts, delayMs }) => {
                console.log(`  ↻ Retrying build (${attempt + 1}/${maxAttempts}) in ${delayMs}ms: ${failure.message}`);
              },
            }
          );
          loopState.buildResult = buildResult;

          if (!buildResult.success) {
            console.log(`  ❌ Build failed: ${buildResult.error}`);
            const finalResult = createBuildFailureResult(loopState.prompt, buildResult);
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

          console.log(`  ✅ Built`);
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
            console.log(`  ⚠️ Artifact check failed: ${artifactStatus.message}`);
            const failure = createFailureDetails(artifactStatus.message, {
              fallbackCategory: artifactStatus.category,
            });
            const finalResult = createArtifactMissingResult(loopState.prompt, failure);
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
          console.log(`  ⚖️ Judging with ${opts.judgeModel}${opts.judgeReasoningEffort ? ` (${opts.judgeReasoningEffort})` : ""}...`);
          const judgeResult = await executeWithRetry(
            () => runJudge(loopState.prompt, loopState.artifacts.pngPath, opts, context.runtimes.judge),
            {
              ...DEFAULT_RETRY_POLICY,
              isSuccess: (result) => Boolean(result?.success),
              classifyFailure: (result) => result?.failure || createFailureDetails(result?.error, {
                fallbackCategory: result?.errorCategory,
              }),
              onRetry: async ({ failure, attempt, maxAttempts, delayMs }) => {
                console.log(`  ↻ Retrying judge (${attempt + 1}/${maxAttempts}) in ${delayMs}ms: ${failure.message}`);
              },
            }
          );
          loopState.judgeResult = judgeResult;

          if (!judgeResult.success) {
            console.log(`  ❌ Judge failed: ${judgeResult.error}`);
            const finalResult = createJudgeFailureResult(loopState.prompt, judgeResult);
            loopState.finalResult = finalResult;

            return {
              status: ORCHESTRATOR_STEP_STATUS.failure,
              terminate: true,
              finalResult,
              terminationReason: "judge-failure",
              failure: judgeResult.failure,
              retry: judgeResult.retry,
            };
          }

          const gaps = judgeResult.parsed.gaps.join("; ");
          console.log(`  📊 Score: ${judgeResult.parsed.totalScore ?? "?"}/14${gaps ? ` | Gap: ${gaps.slice(0, 60)}...` : ""}`);
          if (judgeResult.validation?.status !== "strict") {
            console.log(`  ⚠️  Judge contract ${judgeResult.validation?.status || "invalid"}${judgeResult.validation?.failure?.category ? ` (${judgeResult.validation.failure.category})` : ""}`);
          }

          const finalResult = createSuccessResult(loopState.prompt, judgeResult, loopState.artifacts.pngPath);
          loopState.finalResult = finalResult;

          return {
            status: ORCHESTRATOR_STEP_STATUS.success,
            finalResult,
          };
        },
      },
      onLoopStart: async (loopState) => {
        loopState.metadata.persistence = {
          skillSnapshotBefore: await captureSkillSnapshot({
            baseDir: SKILLS_DIR,
            files: SKILL_FILES,
            label: "builder-skills",
          }),
        };
        console.log(`[${loopState.loopNumber}/${selected.length}] ${loopState.prompt.id} (${loopState.prompt.category})`);
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
      onLoopComplete: async (loopState) => {
        const finalResult = loopState.finalResult || {};
        const isBuildFailure = finalResult.build_success === false;
        const isArtifactFailure = isArtifactFailureCategory(finalResult.error_category);
        const isJudgeFailure = finalResult.build_success === true && finalResult.judge_success === false && !isArtifactFailure;

        await ledger.writeLoopRecord({
          loopNumber: loopState.loopNumber,
          prompt: {
            id: loopState.prompt.id,
            text: loopState.prompt.prompt,
            category: loopState.prompt.category,
            archetype: loopState.prompt.category,
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
            recovery: loopState.recovery,
          },
          builder: {
            request: loopState.buildResult?.requestEnvelope || null,
            text: loopState.buildResult?.response || null,
            errorMessage: loopState.buildResult?.error || null,
            errorCategory: loopState.buildResult?.errorCategory || null,
            retry: loopState.buildResult?.retry || null,
          },
          judge: {
            request: loopState.judgeResult?.requestEnvelope || null,
            parsed: loopState.judgeResult?.parsed || null,
            summary: loopState.judgeResult?.parsed?.summary || null,
            raw: loopState.judgeResult?.raw || null,
            errorMessage: loopState.judgeResult?.error || null,
            errorCategory: loopState.judgeResult?.errorCategory || null,
            validation: loopState.judgeResult?.validation || null,
            retry: loopState.judgeResult?.retry || null,
          },
          improvement: { attempted: false },
          artifacts: [
            { role: "png", kind: "image", path: loopState.artifacts.pngPath },
            { role: "pptx", kind: "presentation", path: loopState.artifacts.pptxPath },
          ],
          skills: {
            beforeLoop: loopState.metadata.persistence?.skillSnapshotBefore || null,
          },
          metadata: {
            mode: opts.mode,
            totalPrompts: selected.length,
          },
        });

        return finalResult;
      },
    });

    const reportPath = join(ledger.runDir, "run-report.json");
    const { report } = await generateRunReportFromPersistence({
      manifestPath: ledger.manifestPath,
      transactionsPath: ledger.transactionsPath,
      outputPath: reportPath,
    });

    console.log(`\n${formatRunReportForConsole(report)}`);

    const scored = results.filter((result) => result?.score != null);
    const allGaps = scored.filter((result) => result.gaps).map((result) => ({
      prompt_id: result.prompt_id,
      gap: result.gaps,
    }));
    if (allGaps.length > 0) {
      console.log(`\n🔍 Gaps found (${allGaps.length}):`);
      allGaps.forEach((gap) => console.log(`   [${gap.prompt_id}] ${gap.gap.slice(0, 80)}`));
    }

    const timestamp = new Date().toISOString().slice(0, 19).replace(/[T:]/g, "-");
    const outFile = join(modeResultsDir, `${runName}-${timestamp}.json`);
    writeFileSync(outFile, JSON.stringify({
      mode: opts.mode,
      opts,
      summary: {
        ...report.summary,
        mode: opts.mode,
        count: selected.length,
        scored: report.summary.scoredLoops,
        avg_score: report.summary.averageScore,
        build_fails: report.summary.buildFailures,
        artifact_fails: report.summary.artifactFailures,
        judge_fails: report.summary.judgeFailures,
      },
      report,
      results,
    }, null, 2));
    await ledger.finalize({
      summary: {
        ...report.summary,
        mode: opts.mode,
        count: selected.length,
        scored: report.summary.scoredLoops,
        avgScore: report.summary.averageScore,
        buildFails: report.summary.buildFailures,
        artifactFails: report.summary.artifactFailures,
        judgeFails: report.summary.judgeFailures,
      },
      finalArtifacts: [
        { role: "results-json", kind: "record", path: outFile },
        { role: "run-report", kind: "record", path: reportPath },
      ],
    });
    console.log(`\n💾 Results saved to: ${outFile}`);
  } finally {
    await destroyRuntime(builderRuntime);
    await destroyRuntime(judgeRuntime);
  }
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
