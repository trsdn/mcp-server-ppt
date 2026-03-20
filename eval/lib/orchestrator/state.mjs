export const ORCHESTRATOR_PHASES = Object.freeze({
  pending: "pending",
  cleanup: "cleanup",
  build: "build",
  verifyArtifact: "verify-artifact",
  judge: "judge",
  improve: "improve",
  recovery: "recovery",
  complete: "complete",
});

export const ORCHESTRATOR_STEP_STATUS = Object.freeze({
  pending: "pending",
  success: "success",
  failure: "failure",
  skipped: "skipped",
});

export const ORCHESTRATOR_STEPS = Object.freeze({
  cleanup: "cleanup",
  build: "build",
  verifyArtifact: "verify-artifact",
  judge: "judge",
  improve: "improve",
});

function cloneStructuredValue(value) {
  if (value == null) return value;
  return JSON.parse(JSON.stringify(value));
}

export function createStepSequence({ includeJudge = true, includeImprove = false } = {}) {
  const sequence = [
    ORCHESTRATOR_STEPS.cleanup,
    ORCHESTRATOR_STEPS.build,
    ORCHESTRATOR_STEPS.verifyArtifact,
  ];

  if (includeJudge) {
    sequence.push(ORCHESTRATOR_STEPS.judge);
  }

  if (includeImprove) {
    sequence.push(ORCHESTRATOR_STEPS.improve);
  }

  return sequence;
}

export function createOrchestrationContext({
  builderRuntime = null,
  judgeRuntime = null,
  improverRuntime = null,
  builderCarryover = [],
  reviewerCarryover = [],
  scoreHistory = [],
  metadata = {},
} = {}) {
  return {
    runtimes: {
      builder: builderRuntime,
      judge: judgeRuntime,
      improver: improverRuntime,
    },
    carryover: {
      builder: cloneStructuredValue(builderCarryover) || [],
      reviewer: cloneStructuredValue(reviewerCarryover) || [],
    },
    scoreHistory: [...scoreHistory],
    metadata: { ...metadata },
    startedAt: new Date().toISOString(),
  };
}

export function createLoopState({
  loopNumber,
  prompt,
  pngPath,
  pptxPath,
  sequence = [],
  carryoverSnapshot = {},
  metadata = {},
} = {}) {
  const stepSequence = [...sequence];

  return {
    loopNumber,
    prompt,
    artifacts: {
      pngPath,
      pptxPath,
    },
    stepSequence,
    stepIndex: 0,
    currentStep: stepSequence[0] || null,
    status: "running",
    phase: ORCHESTRATOR_PHASES.pending,
    stepHistory: [],
    stepResults: {},
    errors: [],
    recovery: {
      needed: false,
      reason: null,
      category: null,
      disposition: null,
      retryable: false,
      attempts: 0,
    },
    carryoverSnapshot: {
      builder: cloneStructuredValue(carryoverSnapshot.builder) || [],
      reviewer: cloneStructuredValue(carryoverSnapshot.reviewer) || [],
    },
    metadata: { ...metadata },
    finalResult: null,
    startedAt: new Date().toISOString(),
    completedAt: null,
  };
}

export function getCurrentStep(loopState) {
  return loopState.currentStep;
}

export function setLoopPhase(loopState, phase) {
  loopState.phase = phase;
  return loopState;
}

export function recordStepTransition(loopState, step, status, details = {}) {
  const entry = {
    step,
    status,
    phase: loopState.phase,
    timestamp: new Date().toISOString(),
    ...details,
  };

  loopState.stepHistory.push(entry);
  loopState.stepResults[step] = entry;
  return entry;
}

export function advanceLoopState(loopState, nextStep = undefined) {
  if (typeof nextStep === "string") {
    const nextIndex = loopState.stepSequence.indexOf(nextStep);
    loopState.stepIndex = nextIndex >= 0 ? nextIndex : loopState.stepSequence.length;
  } else {
    loopState.stepIndex += 1;
  }

  loopState.currentStep = loopState.stepSequence[loopState.stepIndex] || null;

  if (!loopState.currentStep) {
    loopState.phase = ORCHESTRATOR_PHASES.complete;
    if (loopState.status === "running") {
      loopState.status = "completed";
    }
    loopState.completedAt = new Date().toISOString();
  }

  return loopState.currentStep;
}

export function finalizeLoopState(loopState, { status = "completed", finalResult = loopState.finalResult } = {}) {
  loopState.status = status;
  loopState.phase = status === "failed" ? ORCHESTRATOR_PHASES.recovery : ORCHESTRATOR_PHASES.complete;
  loopState.currentStep = null;
  loopState.stepIndex = loopState.stepSequence.length;
  loopState.finalResult = finalResult;
  loopState.completedAt = new Date().toISOString();
  return loopState;
}

export function pushRollingItem(items, value, limit = 5) {
  if (!value) return items;

  items.push(cloneStructuredValue(value));
  if (items.length > limit) {
    items.splice(0, items.length - limit);
  }

  return items;
}
