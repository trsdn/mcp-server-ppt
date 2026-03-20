import { join, normalize, sep } from "path";

export const EVAL_MODES = Object.freeze({
  baseline: "baseline",
  tuning: "tuning",
});

function normalizeEvalModeValue(value) {
  const normalized = String(value ?? "").trim().toLowerCase();
  return Object.values(EVAL_MODES).includes(normalized) ? normalized : null;
}

export function resolveEvalMode({
  configuredMode = null,
  hasImprover = false,
  defaultMode = EVAL_MODES.baseline,
} = {}) {
  const normalizedConfiguredMode = normalizeEvalModeValue(configuredMode);
  if (normalizedConfiguredMode) {
    return normalizedConfiguredMode;
  }

  return hasImprover ? EVAL_MODES.tuning : defaultMode;
}

export function ensureEvalMode(mode, source = "mode") {
  const normalizedMode = normalizeEvalModeValue(mode);
  if (!normalizedMode) {
    throw new Error(`${source} must be one of: ${Object.values(EVAL_MODES).join(", ")}`);
  }

  return normalizedMode;
}

export function getModeScopedDirectory(baseDir, mode) {
  const normalizedMode = ensureEvalMode(mode);
  const normalizedBaseDir = normalize(baseDir);
  return normalizedBaseDir.endsWith(`${sep}${normalizedMode}`) || normalizedBaseDir === normalizedMode
    ? normalizedBaseDir
    : join(normalizedBaseDir, normalizedMode);
}

export function getModeTaggedName(name, mode) {
  const normalizedMode = ensureEvalMode(mode);
  const trimmedName = String(name ?? "").trim();
  return trimmedName.endsWith(`-${normalizedMode}`)
    ? trimmedName
    : `${trimmedName || "eval-run"}-${normalizedMode}`;
}
