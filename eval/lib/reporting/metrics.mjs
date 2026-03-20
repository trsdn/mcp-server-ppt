import { readFile, writeFile } from "fs/promises";
import { dirname, join, resolve } from "path";
import { JUDGE_DIMENSION_KEYS } from "../protocol/index.mjs";
import {
  readLoopRecord,
  readRunManifest,
  readTransactionLedger,
  stableJsonStringify,
} from "../persistence/index.mjs";

const REPORT_CONTRACT = "eval-run-report/v1";

function roundNumber(value, digits = 2) {
  if (!Number.isFinite(value)) return null;
  return Number.parseFloat(value.toFixed(digits));
}

function average(values) {
  if (!values.length) return null;
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function median(values) {
  if (!values.length) return null;
  const sorted = [...values].sort((left, right) => left - right);
  const midpoint = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0
    ? (sorted[midpoint - 1] + sorted[midpoint]) / 2
    : sorted[midpoint];
}

function countBy(items, selector) {
  const counts = {};
  for (const item of items) {
    const key = selector(item);
    if (!key) continue;
    counts[key] = (counts[key] || 0) + 1;
  }
  return counts;
}

function normalizeGapText(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function inferExpectedLoops(manifest) {
  const configuredLoops = manifest?.metadata?.config?.loops;
  if (Number.isInteger(configuredLoops)) return configuredLoops;

  const selectedPromptIds = manifest?.metadata?.selectedPromptIds;
  if (Array.isArray(selectedPromptIds)) return selectedPromptIds.length;

  return manifest?.counters?.recordedLoops ?? 0;
}

function getLoopPromptId(loopRecord) {
  return loopRecord?.loop?.prompt?.id || null;
}

function getLoopNumber(loopRecord) {
  return loopRecord?.loop?.number ?? null;
}

function getScoredLoops(loopRecords) {
  return loopRecords.filter((loop) => Number.isFinite(loop?.outcome?.score));
}

function getDimensionPoints(loopRecords, dimensionKey) {
  return getScoredLoops(loopRecords)
    .map((loop) => {
      const score = loop?.judge?.response?.parsed?.dimensionScores?.[dimensionKey]?.score;
      return Number.isFinite(score)
        ? {
          loopNumber: getLoopNumber(loop),
          promptId: getLoopPromptId(loop),
          score,
        }
        : null;
    })
    .filter(Boolean);
}

function computeWindowTrend(values, windowSize = 3) {
  if (values.length < windowSize * 2) {
    return {
      windowSize,
      previousAverage: null,
      recentAverage: null,
      delta: null,
      regressionDetected: false,
    };
  }

  const previousWindow = values.slice(-windowSize * 2, -windowSize);
  const recentWindow = values.slice(-windowSize);
  const previousAverage = average(previousWindow);
  const recentAverage = average(recentWindow);
  const delta = recentAverage - previousAverage;

  return {
    windowSize,
    previousAverage: roundNumber(previousAverage),
    recentAverage: roundNumber(recentAverage),
    delta: roundNumber(delta),
    regressionDetected: delta <= -0.5,
  };
}

function computeDirection(delta) {
  if (!Number.isFinite(delta) || Math.abs(delta) < 0.01) return "flat";
  return delta > 0 ? "improving" : "regressing";
}

function computeScoreSummary(loopRecords) {
  const scoredLoops = getScoredLoops(loopRecords);
  const values = scoredLoops.map((loop) => loop.outcome.score);
  const points = scoredLoops.map((loop) => ({
    loopNumber: getLoopNumber(loop),
    promptId: getLoopPromptId(loop),
    score: loop.outcome.score,
    maxScore: loop.outcome.maxScore,
  }));
  const trendWindow = computeWindowTrend(values);
  const firstScore = values[0] ?? null;
  const lastScore = values.length ? values[values.length - 1] : null;
  const deltaFirstToLast = Number.isFinite(firstScore) && Number.isFinite(lastScore)
    ? lastScore - firstScore
    : null;

  return {
    scoredLoops: scoredLoops.length,
    unscoredLoops: loopRecords.length - scoredLoops.length,
    averageScore: roundNumber(average(values)),
    medianScore: roundNumber(median(values)),
    minScore: values.length ? Math.min(...values) : null,
    maxScore: values.length ? Math.max(...values) : null,
    firstScore,
    lastScore,
    deltaFirstToLast: roundNumber(deltaFirstToLast),
    direction: computeDirection(deltaFirstToLast),
    trendWindow,
    regressionDetected: Boolean(trendWindow.regressionDetected),
    history: points,
  };
}

function computeDimensionSummary(loopRecords) {
  return Object.fromEntries(
    JUDGE_DIMENSION_KEYS.map((dimensionKey) => {
      const points = getDimensionPoints(loopRecords, dimensionKey);
      const values = points.map((point) => point.score);
      const trendWindow = computeWindowTrend(values);
      const firstScore = values[0] ?? null;
      const lastScore = values.length ? values[values.length - 1] : null;
      const deltaFirstToLast = Number.isFinite(firstScore) && Number.isFinite(lastScore)
        ? lastScore - firstScore
        : null;

      return [
        dimensionKey,
        {
          averageScore: roundNumber(average(values)),
          minScore: values.length ? Math.min(...values) : null,
          maxScore: values.length ? Math.max(...values) : null,
          firstScore,
          lastScore,
          deltaFirstToLast: roundNumber(deltaFirstToLast),
          direction: computeDirection(deltaFirstToLast),
          trendWindow,
          regressionDetected: Boolean(trendWindow.regressionDetected),
          points,
        },
      ];
    })
  );
}

function extractGapItems(loopRecord) {
  const gaps = loopRecord?.judge?.response?.parsed?.gaps;
  if (Array.isArray(gaps)) {
    return gaps.filter(Boolean);
  }

  const rawGaps = loopRecord?.outcome?.errorMessage;
  return typeof rawGaps === "string" && rawGaps.includes(";")
    ? rawGaps.split(";").map((gap) => gap.trim()).filter(Boolean)
    : [];
}

function buildGapSummary(loopRecords) {
  const gapMap = new Map();

  for (const loop of loopRecords) {
    const loopNumber = getLoopNumber(loop);
    const promptId = getLoopPromptId(loop);
    for (const gap of extractGapItems(loop)) {
      const normalizedGap = normalizeGapText(gap);
      if (!normalizedGap) continue;

      if (!gapMap.has(normalizedGap)) {
        gapMap.set(normalizedGap, {
          gap,
          normalizedGap,
          count: 0,
          loopNumbers: [],
          promptIds: new Set(),
        });
      }

      const entry = gapMap.get(normalizedGap);
      entry.count += 1;
      entry.loopNumbers.push(loopNumber);
      entry.promptIds.add(promptId);
    }
  }

  const repeated = [...gapMap.values()]
    .map((entry) => {
      const sortedLoops = [...entry.loopNumbers].filter(Number.isFinite).sort((left, right) => left - right);
      let longestStreak = 0;
      let currentStreak = 0;
      let previousLoop = null;

      for (const loopNumber of sortedLoops) {
        if (previousLoop != null && loopNumber === previousLoop + 1) {
          currentStreak += 1;
        } else {
          currentStreak = 1;
        }
        longestStreak = Math.max(longestStreak, currentStreak);
        previousLoop = loopNumber;
      }

      return {
        gap: entry.gap,
        normalizedGap: entry.normalizedGap,
        count: entry.count,
        loopNumbers: sortedLoops,
        promptIds: [...entry.promptIds].filter(Boolean).sort(),
        longestStreak,
        repeatedDetected: entry.count >= 2,
      };
    })
    .sort((left, right) => (
      right.count - left.count
      || right.longestStreak - left.longestStreak
      || left.normalizedGap.localeCompare(right.normalizedGap)
    ));

  return {
    uniqueGapCount: gapMap.size,
    repeatedGapCount: repeated.filter((entry) => entry.repeatedDetected).length,
    repeatedGaps: repeated.filter((entry) => entry.repeatedDetected),
    topGaps: repeated.slice(0, 10),
  };
}

function buildRetrySummary(loopRecords) {
  const byStep = {
    build: { loopsRetried: 0, extraAttempts: 0 },
    judge: { loopsRetried: 0, extraAttempts: 0 },
    improvement: { loopsRetried: 0, extraAttempts: 0 },
  };

  let recoveredLoops = 0;
  let loopsWithRetries = 0;

  for (const loop of loopRecords) {
    const retries = loop?.retries || {};
    const extraBuildAttempts = retries.build || 0;
    const extraJudgeAttempts = retries.judge || 0;
    const extraImprovementAttempts = retries.improvement || 0;
    const totalExtraAttempts = extraBuildAttempts + extraJudgeAttempts + extraImprovementAttempts;

    if (extraBuildAttempts > 0) {
      byStep.build.loopsRetried += 1;
      byStep.build.extraAttempts += extraBuildAttempts;
    }
    if (extraJudgeAttempts > 0) {
      byStep.judge.loopsRetried += 1;
      byStep.judge.extraAttempts += extraJudgeAttempts;
    }
    if (extraImprovementAttempts > 0) {
      byStep.improvement.loopsRetried += 1;
      byStep.improvement.extraAttempts += extraImprovementAttempts;
    }

    if (totalExtraAttempts > 0) {
      loopsWithRetries += 1;
      if (loop?.outcome?.status === "completed") {
        recoveredLoops += 1;
      }
    }
  }

  return {
    loopsWithRetries,
    recoveredLoops,
    byStep,
    totalExtraAttempts: byStep.build.extraAttempts + byStep.judge.extraAttempts + byStep.improvement.extraAttempts,
  };
}

function buildFailureSummary(loopRecords, transactions) {
  const failedLoops = loopRecords.filter((loop) => loop?.outcome?.status !== "completed");
  const failureCategoryCounts = countBy(failedLoops, (loop) => loop?.outcome?.errorCategory);
  const dispositions = countBy(failedLoops, (loop) => loop?.diagnostics?.recovery?.disposition);
  const retryableFailures = failedLoops.filter((loop) => loop?.diagnostics?.recovery?.retryable === true).length;
  const fatalFailures = failedLoops.length - retryableFailures;
  const transactionFailureCategoryCounts = countBy(
    transactions.filter((entry) => entry?.status !== "completed"),
    (entry) => entry?.errorCategory
  );

  return {
    failedLoops: failedLoops.length,
    failureCategoryCounts,
    transactionFailureCategoryCounts,
    dispositions,
    retryableFailures,
    fatalFailures,
  };
}

function buildOutcomeSummary(loopRecords, expectedLoops) {
  const statusCounts = countBy(loopRecords, (loop) => loop?.outcome?.status);
  const buildStatusCounts = countBy(loopRecords, (loop) => loop?.outcome?.buildStatus);
  const judgeStatusCounts = countBy(loopRecords, (loop) => loop?.outcome?.judgeStatus);
  const completedLoops = statusCounts.completed || 0;
  const scoredLoops = getScoredLoops(loopRecords).length;

  return {
    expectedLoops,
    recordedLoops: loopRecords.length,
    statusCounts,
    buildStatusCounts,
    judgeStatusCounts,
    completedLoops,
    scoredLoops,
    completionRate: expectedLoops > 0 ? roundNumber(completedLoops / expectedLoops, 4) : null,
    scoredRate: expectedLoops > 0 ? roundNumber(scoredLoops / expectedLoops, 4) : null,
  };
}

function buildOperatorSummary({
  manifest,
  outcomeSummary,
  scoreSummary,
  failureSummary,
  retrySummary,
  dimensionSummary,
  gapSummary,
}) {
  const alerts = [];

  if (scoreSummary.regressionDetected) {
    alerts.push(`overall score regressed by ${scoreSummary.trendWindow.delta} over the most recent ${scoreSummary.trendWindow.windowSize} scored loops`);
  }

  for (const [dimensionKey, summary] of Object.entries(dimensionSummary)) {
    if (summary.regressionDetected) {
      alerts.push(`${dimensionKey} regressed by ${summary.trendWindow.delta} over the most recent ${summary.trendWindow.windowSize} scored loops`);
    }
  }

  const topRepeatedGap = gapSummary.repeatedGaps[0];
  if (topRepeatedGap) {
    alerts.push(`repeated gap "${topRepeatedGap.gap}" appeared ${topRepeatedGap.count} times`);
  }

  const topFailureCategory = Object.entries(failureSummary.failureCategoryCounts)
    .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0]))[0];
  if (topFailureCategory) {
    alerts.push(`top failure category ${topFailureCategory[0]} occurred ${topFailureCategory[1]} times`);
  }

  const averageText = scoreSummary.averageScore == null ? "n/a" : `${scoreSummary.averageScore}`;
  const headline = `${manifest.runName}: ${outcomeSummary.scoredLoops}/${outcomeSummary.expectedLoops} scored, avg ${averageText}`;
  const lines = [
    `Run ${manifest.runId} (${manifest.runner})`,
    `Scored ${outcomeSummary.scoredLoops}/${outcomeSummary.expectedLoops}; completed ${outcomeSummary.completedLoops}/${outcomeSummary.expectedLoops}`,
    `Failures: ${failureSummary.failedLoops} total; retries on ${retrySummary.loopsWithRetries} loops (${retrySummary.totalExtraAttempts} extra attempts)`,
    `Repeated gaps: ${gapSummary.repeatedGapCount}; unique failure categories: ${Object.keys(failureSummary.failureCategoryCounts).length}`,
  ];

  return {
    headline,
    lines,
    alerts,
  };
}

export async function loadRunDataFromPersistence({
  manifestPath,
  transactionsPath = null,
}) {
  const resolvedManifestPath = resolve(manifestPath);
  const manifest = await readRunManifest(resolvedManifestPath);
  const runDir = dirname(resolvedManifestPath);
  const resolvedTransactionsPath = resolve(transactionsPath || manifest?.directories?.transactionsPath || join(runDir, "transactions.jsonl"));
  let transactions = [];
  try {
    await readFile(resolvedTransactionsPath, "utf8");
    transactions = await readTransactionLedger(resolvedTransactionsPath);
  } catch {
    transactions = [];
  }
  const loopIndex = Array.isArray(manifest?.loopIndex) ? manifest.loopIndex : [];
  const loopRecords = await Promise.all(
    loopIndex.map(async (entry) => readLoopRecord(join(runDir, entry.recordPath)))
  );

  loopRecords.sort((left, right) => (getLoopNumber(left) || 0) - (getLoopNumber(right) || 0));

  return {
    manifest,
    loopRecords,
    transactions,
    manifestPath: resolvedManifestPath,
    transactionsPath: resolvedTransactionsPath,
    runDir,
  };
}

export function createRunReport({
  manifest,
  loopRecords,
  transactions,
  manifestPath,
  transactionsPath,
  reportGeneratedAt = new Date().toISOString(),
}) {
  const expectedLoops = inferExpectedLoops(manifest);
  const outcomeSummary = buildOutcomeSummary(loopRecords, expectedLoops);
  const scoreSummary = computeScoreSummary(loopRecords);
  const dimensionSummary = computeDimensionSummary(loopRecords);
  const failureSummary = buildFailureSummary(loopRecords, transactions);
  const retrySummary = buildRetrySummary(loopRecords);
  const gapSummary = buildGapSummary(loopRecords);
  const operatorSummary = buildOperatorSummary({
    manifest,
    outcomeSummary,
    scoreSummary,
    failureSummary,
    retrySummary,
    dimensionSummary,
    gapSummary,
  });

  return {
    contract: REPORT_CONTRACT,
    generatedAt: reportGeneratedAt,
    run: {
      runId: manifest?.runId || null,
      runName: manifest?.runName || null,
      runner: manifest?.runner || null,
      manifestPath,
      transactionsPath,
      startedAt: manifest?.startedAt || null,
      metadata: manifest?.metadata || {},
      counters: manifest?.counters || {},
    },
    consistency: {
      loopIndexCount: Array.isArray(manifest?.loopIndex) ? manifest.loopIndex.length : 0,
      loopRecordCount: loopRecords.length,
      transactionCount: transactions.length,
      transactionLoopMismatch: transactions.length !== loopRecords.length,
    },
    outcomes: outcomeSummary,
    scores: scoreSummary,
    dimensions: dimensionSummary,
    failures: failureSummary,
    retries: retrySummary,
    gaps: gapSummary,
    operatorSummary,
    summary: {
      expectedLoops: outcomeSummary.expectedLoops,
      recordedLoops: outcomeSummary.recordedLoops,
      completedLoops: outcomeSummary.completedLoops,
      scoredLoops: outcomeSummary.scoredLoops,
      averageScore: scoreSummary.averageScore,
      avgScore: scoreSummary.averageScore,
      buildFailures: outcomeSummary.buildStatusCounts.failed || 0,
      artifactFailures: (failureSummary.failureCategoryCounts.artifact_missing || 0) + (failureSummary.failureCategoryCounts.artifact_invalid || 0),
      judgeFailures: outcomeSummary.judgeStatusCounts.failed || 0,
      regressionDetected: scoreSummary.regressionDetected || Object.values(dimensionSummary).some((summary) => summary.regressionDetected),
      repeatedGapCount: gapSummary.repeatedGapCount,
      topFailureCategory: Object.entries(failureSummary.failureCategoryCounts)
        .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0]))[0]?.[0] || null,
      totalExtraAttempts: retrySummary.totalExtraAttempts,
    },
  };
}

export async function writeRunReport(outputPath, report) {
  await writeFile(outputPath, `${stableJsonStringify(report)}\n`, "utf8");
  return outputPath;
}

export async function generateRunReportFromPersistence({
  manifestPath,
  transactionsPath = null,
  outputPath = null,
}) {
  const data = await loadRunDataFromPersistence({ manifestPath, transactionsPath });
  const report = createRunReport(data);

  if (outputPath) {
    await writeRunReport(outputPath, report);
  }

  return {
    ...data,
    report,
    outputPath,
  };
}

export function formatRunReportForConsole(report) {
  const lines = [];
  lines.push(report.operatorSummary.headline);
  for (const line of report.operatorSummary.lines) {
    lines.push(`  ${line}`);
  }
  if (report.operatorSummary.alerts.length > 0) {
    lines.push("  Alerts:");
    for (const alert of report.operatorSummary.alerts.slice(0, 5)) {
      lines.push(`    - ${alert}`);
    }
  }
  return lines.join("\n");
}
