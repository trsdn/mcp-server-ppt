import {
  ORCHESTRATOR_PHASES,
  ORCHESTRATOR_STEP_STATUS,
  ORCHESTRATOR_STEPS,
  advanceLoopState,
  finalizeLoopState,
  getCurrentStep,
  recordStepTransition,
  setLoopPhase,
} from "./state.mjs";

function resolvePhaseForStep(step) {
  switch (step) {
    case ORCHESTRATOR_STEPS.cleanup:
      return ORCHESTRATOR_PHASES.cleanup;
    case ORCHESTRATOR_STEPS.build:
      return ORCHESTRATOR_PHASES.build;
    case ORCHESTRATOR_STEPS.verifyArtifact:
      return ORCHESTRATOR_PHASES.verifyArtifact;
    case ORCHESTRATOR_STEPS.judge:
      return ORCHESTRATOR_PHASES.judge;
    case ORCHESTRATOR_STEPS.improve:
      return ORCHESTRATOR_PHASES.improve;
    default:
      return ORCHESTRATOR_PHASES.pending;
  }
}

function sanitizeOutcome(outcome = {}) {
  const {
    status,
    terminate,
    nextStep,
    finalStatus,
    finalResult,
    phase,
    ...details
  } = outcome;

  return details;
}

function resolveFailureContext(outcome = {}) {
  if (!outcome || outcome.status !== ORCHESTRATOR_STEP_STATUS.failure) {
    return null;
  }

  const retry = outcome.retry || {};
  const failure = outcome.failure || {};
  const reason = failure.message || outcome.error || outcome.terminationReason || "Step failed";

  return {
    needed: true,
    reason,
    category: failure.category || outcome.failureCategory || null,
    disposition: failure.disposition || outcome.failureDisposition || null,
    retryable: Boolean(
      failure.retryable
      ?? outcome.retryable
      ?? retry.exhausted
      ?? false
    ),
    attempts: retry.attempts || outcome.attempts || 1,
  };
}

export async function runLoopOrchestrator({
  items = [],
  context,
  createLoopState,
  handlers,
  shouldSkipStep,
  onLoopStart,
  onStepStart,
  onStepComplete,
  onLoopComplete,
  continueOnUnhandledStepError = false,
} = {}) {
  const results = [];

  for (let index = 0; index < items.length; index++) {
    const loopState = createLoopState({ item: items[index], index, context });

    if (!loopState) {
      throw new Error("createLoopState must return a loop state object.");
    }

    await onLoopStart?.(loopState, context);

    while (getCurrentStep(loopState)) {
      const step = getCurrentStep(loopState);
      const handler = handlers?.[step];

      if (!handler) {
        throw new Error(`No orchestrator handler registered for step '${step}'.`);
      }

      setLoopPhase(loopState, resolvePhaseForStep(step));

      if (shouldSkipStep?.(step, loopState, context) === true) {
        recordStepTransition(loopState, step, ORCHESTRATOR_STEP_STATUS.skipped);
        await onStepComplete?.(step, loopState, context, { status: ORCHESTRATOR_STEP_STATUS.skipped });
        advanceLoopState(loopState);
        continue;
      }

      await onStepStart?.(step, loopState, context);

      let outcome;
      try {
        outcome = await handler(loopState, context);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error || "Unknown step error");
        loopState.errors.push({ step, message });
        loopState.recovery = { needed: true, reason: message };
        setLoopPhase(loopState, ORCHESTRATOR_PHASES.recovery);
        recordStepTransition(loopState, step, ORCHESTRATOR_STEP_STATUS.failure, { error: message });
        finalizeLoopState(loopState, { status: "failed", finalResult: loopState.finalResult });

        if (!continueOnUnhandledStepError) {
          throw error;
        }

        break;
      }

      const status = outcome?.status || ORCHESTRATOR_STEP_STATUS.success;
      const failureContext = resolveFailureContext({
        ...outcome,
        status,
      });

      if (outcome?.phase) {
        setLoopPhase(loopState, outcome.phase);
      }

      if (failureContext) {
        loopState.recovery = failureContext;
        if (!outcome?.phase) {
          setLoopPhase(loopState, ORCHESTRATOR_PHASES.recovery);
        }
      }

      if (outcome?.finalResult !== undefined) {
        loopState.finalResult = outcome.finalResult;
      }

      recordStepTransition(loopState, step, status, sanitizeOutcome(outcome));
      await onStepComplete?.(step, loopState, context, outcome);

      if (outcome?.terminate) {
        finalizeLoopState(loopState, {
          status: outcome.finalStatus || "completed",
          finalResult: outcome.finalResult ?? loopState.finalResult,
        });
        break;
      }

      advanceLoopState(loopState, outcome?.nextStep);
    }

    if (!loopState.completedAt) {
      finalizeLoopState(loopState, {
        status: loopState.status === "running" ? "completed" : loopState.status,
        finalResult: loopState.finalResult,
      });
    }

    const completedResult = await onLoopComplete?.(loopState, context);
    results.push(completedResult ?? loopState.finalResult ?? loopState);
  }

  return {
    context,
    results,
  };
}
