import { dirname, isAbsolute, join, resolve } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));

export const RUNTIME_ROOT = __dirname;
export const EVAL_ROOT = join(RUNTIME_ROOT, "..", "..");
export const REPO_ROOT = join(EVAL_ROOT, "..");
export const SKILLS_DIR = join(REPO_ROOT, "skills", "shared");
export const ARCHETYPES_DIR = join(REPO_ROOT, "src", "PptMcp.Core", "Data", "archetypes");
export const CLI_PATH = join(REPO_ROOT, "src", "PptMcp.CLI", "bin", "Release", "net9.0-windows", "pptcli.exe");
export const EVAL_ASSET_REPO_ROOT_ENVIRONMENT_VARIABLE = "PPTMCP_EVAL_ASSET_REPO_ROOT";

export function getEvalAssetRepoRoot() {
  const configuredRoot = process.env[EVAL_ASSET_REPO_ROOT_ENVIRONMENT_VARIABLE];
  return configuredRoot ? resolve(configuredRoot) : REPO_ROOT;
}

export function getEvalAssetEvalRoot() {
  return join(getEvalAssetRepoRoot(), "eval");
}

export function resolveEvalAssetPath(...segments) {
  return join(getEvalAssetEvalRoot(), ...segments);
}

export function resolveEvalAssetPathFromRelative(relativeOrAbsolutePath) {
  const candidate = String(relativeOrAbsolutePath ?? "").trim();
  if (!candidate) {
    throw new Error("Eval asset path must be provided.");
  }

  return isAbsolute(candidate) ? resolve(candidate) : resolveEvalAssetPath(candidate);
}

export const EVAL_INPUT_ROOT = resolveEvalAssetPath("input");
export const EVAL_OUTPUT_ROOT = resolveEvalAssetPath("output");
export const EVAL_RESULTS_ROOT = resolveEvalAssetPath("results");
export const EVAL_DATA_ROOT = resolveEvalAssetPath("data");
export const EVAL_REFERENCE_CATALOG_ROOT = join(EVAL_DATA_ROOT, "archetype-references");

/**
 * Maps eval config archetype IDs (which may be variant names) to their
 * canonical family file basenames in ARCHETYPES_DIR.
 */
const ARCHETYPE_FAMILY_MAP = Object.freeze({
  // Direct matches (archetype ID === family file basename)
  "big-number": "big-number",
  "kpi-card-dashboard": "kpi-card-dashboard",
  "operational-kpi": "operational-kpi",
  "column-bar-chart": "column-bar-chart",
  "chart-insight-callout": "chart-insight-callout",
  "framework": "framework",
  "simple-table": "simple-table",
  "waterfall-chart": "waterfall-chart",
  "comparison": "comparison",
  "timeline-roadmap": "timeline-roadmap",
  "process-diagram": "process-diagram",
  "executive-summary": "executive-summary",
  "recommendations": "recommendations",
  "quote": "quote",
  "map": "map",
  "appendix": "appendix",
  "title-slide": "title-slide",
  // Variant aliases → parent family
  "kpi-dashboard": "kpi-card-dashboard",
  "map-slide": "map",
  "architecture-diagram": "framework",
  "balanced-scorecard": "kpi-card-dashboard",
  "funnel": "process-diagram",
  "maturity-model": "framework",
  "stakeholder-map": "framework",
  "strategic-pillars": "framework",
  "swot-analysis": "framework",
});

/**
 * Resolves an archetype ID from an eval config to its canonical family ID.
 * Returns the input unchanged if no mapping exists.
 */
export function resolveArchetypeFamily(archetypeId) {
  return ARCHETYPE_FAMILY_MAP[archetypeId] || archetypeId;
}
export const MCP_SERVER_PATH = join(REPO_ROOT, "src", "PptMcp.McpServer", "bin", "Release", "net9.0-windows", "PptMcp.McpServer.exe");
export const ISOLATED_WORKER_PATH = join(EVAL_ROOT, "copilot-isolated-worker.mjs");
