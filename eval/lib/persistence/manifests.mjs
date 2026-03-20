import { basename, extname, relative, resolve } from "path";
import { hashString, getFileDigest } from "./hashing.mjs";
import { PERSISTENCE_CONTRACTS } from "./records.mjs";

function compactText(value, maxLength = 4000) {
  return typeof value === "string" ? value.trim().slice(0, maxLength) : null;
}

function resolvePath(pathValue, { baseDir = null } = {}) {
  if (!pathValue) return null;
  return baseDir ? resolve(baseDir, pathValue) : resolve(pathValue);
}

function detectArtifactKind(pathValue, fallbackKind = "file") {
  const extension = extname(pathValue || "").toLowerCase();
  switch (extension) {
    case ".png":
    case ".jpg":
    case ".jpeg":
    case ".webp":
      return "image";
    case ".pptx":
    case ".pptm":
      return "presentation";
    case ".json":
    case ".jsonl":
      return "record";
    default:
      return fallbackKind;
  }
}

export async function createArtifactManifest({
  artifacts = [],
  relativeTo = null,
  baseDir = null,
  capturedAt = new Date().toISOString(),
}) {
  const relativeRoot = relativeTo ? resolve(relativeTo) : null;

  const items = await Promise.all(
    artifacts.map(async (artifact, index) => {
      const absolutePath = resolvePath(artifact?.path, { baseDir });
      const digest = absolutePath
        ? await getFileDigest(absolutePath)
        : { exists: false, bytes: null, modifiedAt: null, sha256: null, error: null };

      return {
        artifactId: compactText(artifact?.id ?? "", 160) || `artifact-${index + 1}`,
        role: compactText(artifact?.role ?? "", 160) || `artifact-${index + 1}`,
        label: compactText(artifact?.label ?? "", 240),
        kind: compactText(artifact?.kind ?? "", 160) || detectArtifactKind(absolutePath || artifact?.path, "file"),
        path: absolutePath,
        relativePath: absolutePath && relativeRoot ? relative(relativeRoot, absolutePath) : absolutePath,
        fileName: absolutePath ? basename(absolutePath) : null,
        exists: digest.exists,
        required: artifact?.required !== false,
        bytes: digest.bytes,
        modifiedAt: digest.modifiedAt,
        sha256: digest.sha256,
        error: compactText(digest.error ?? artifact?.error ?? "", 4000),
        metadata: artifact?.metadata ?? {},
      };
    })
  );

  return {
    contract: PERSISTENCE_CONTRACTS.artifactManifest,
    capturedAt,
    itemCount: items.length,
    fingerprintSha256: hashString(
      items.map((item) => `${item.role}:${item.relativePath || item.path || "missing"}:${item.sha256 || "missing"}`).join("\n")
    ),
    items,
  };
}

export async function captureSkillSnapshot({
  files = [],
  baseDir,
  relativeTo = baseDir,
  capturedAt = new Date().toISOString(),
  label = "builder-skills",
}) {
  const uniqueFiles = [...new Set(files.filter(Boolean))];
  const relativeRoot = relativeTo ? resolve(relativeTo) : null;

  const items = await Promise.all(
    uniqueFiles.map(async (file) => {
      const absolutePath = resolve(baseDir, file);
      const digest = await getFileDigest(absolutePath);

      return {
        file,
        label,
        path: absolutePath,
        relativePath: relativeRoot ? relative(relativeRoot, absolutePath) : absolutePath,
        exists: digest.exists,
        bytes: digest.bytes,
        modifiedAt: digest.modifiedAt,
        sha256: digest.sha256,
        error: compactText(digest.error, 4000),
      };
    })
  );

  return {
    contract: PERSISTENCE_CONTRACTS.skillSnapshot,
    label,
    baseDir: resolve(baseDir),
    capturedAt,
    itemCount: items.length,
    fingerprintSha256: hashString(
      items.map((item) => `${item.relativePath}:${item.sha256 || "missing"}:${item.modifiedAt || "missing"}`).join("\n")
    ),
    items,
  };
}
