import { createHash } from "crypto";
import { createReadStream } from "fs";
import { stat } from "fs/promises";

function toErrorMessage(error) {
  return error instanceof Error ? error.message : String(error);
}

export function hashString(value, algorithm = "sha256") {
  const hash = createHash(algorithm);
  hash.update(String(value ?? ""), "utf8");
  return hash.digest("hex");
}

export async function hashFile(filePath, algorithm = "sha256") {
  return new Promise((resolve, reject) => {
    const hash = createHash(algorithm);
    const stream = createReadStream(filePath);

    stream.on("data", (chunk) => hash.update(chunk));
    stream.on("end", () => resolve(hash.digest("hex")));
    stream.on("error", reject);
  });
}

export async function getFileDigest(filePath) {
  try {
    const fileStat = await stat(filePath);
    return {
      exists: true,
      bytes: fileStat.size,
      modifiedAt: fileStat.mtime.toISOString(),
      sha256: await hashFile(filePath),
      error: null,
    };
  } catch (error) {
    if (error && typeof error === "object" && "code" in error && error.code === "ENOENT") {
      return {
        exists: false,
        bytes: null,
        modifiedAt: null,
        sha256: null,
        error: null,
      };
    }

    return {
      exists: false,
      bytes: null,
      modifiedAt: null,
      sha256: null,
      error: toErrorMessage(error),
    };
  }
}
