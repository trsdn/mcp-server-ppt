'use strict';

// Resolves the PptMcp MCP server binary, downloading the self-contained Windows
// build from the matching GitHub release on first use.
//
// Download is lazy rather than a postinstall hook on purpose. npm 12 blocks
// dependency install scripts by default, and `npx` installs this package as a
// dependency of a temporary root, so a postinstall hook is not something we can
// rely on continuing to run. Resolving in `bin` always runs.
//
// Everything this module prints goes to stderr. stdout is the MCP transport.

const fs = require('fs');
const os = require('os');
const path = require('path');
const https = require('https');
const { spawnSync } = require('child_process');

const REPO = 'trsdn/mcp-server-ppt';
const EXE_NAME = 'PptMcp.McpServer.exe';

function packageVersion() {
  return require('../package.json').version;
}

function assetName(version = packageVersion()) {
  return `PptMcp-MCP-Server-${version}-win-x64.zip`;
}

function assetUrl(version = packageVersion()) {
  if (process.env.MCP_SERVER_PPT_ASSET_URL) {
    return process.env.MCP_SERVER_PPT_ASSET_URL;
  }
  return `https://github.com/${REPO}/releases/download/v${version}/${assetName(version)}`;
}

function installRoot(version) {
  const base =
    process.env.MCP_SERVER_PPT_CACHE ||
    path.join(process.env.LOCALAPPDATA || os.tmpdir(), 'mcp-server-ppt');
  return path.join(base, `runtime-${version}`);
}

function log(message) {
  process.stderr.write(`[mcp-server-ppt] ${message}\n`);
}

function download(url, destination, redirectsLeft = 5) {
  return new Promise((resolve, reject) => {
    https
      .get(url, { headers: { 'user-agent': `mcp-server-ppt/${packageVersion()}` } }, (response) => {
        const status = response.statusCode;

        if (status >= 300 && status < 400 && response.headers.location) {
          response.resume();
          if (redirectsLeft === 0) {
            reject(new Error(`Too many redirects while downloading ${url}`));
            return;
          }
          resolve(download(new URL(response.headers.location, url).toString(), destination, redirectsLeft - 1));
          return;
        }

        if (status !== 200) {
          response.resume();
          reject(new Error(`Download failed with HTTP ${status}: ${url}`));
          return;
        }

        const total = Number(response.headers['content-length']) || 0;
        let received = 0;
        let lastReport = 0;

        const file = fs.createWriteStream(destination);
        response.on('data', (chunk) => {
          received += chunk.length;
          const now = Date.now();
          if (total && now - lastReport > 2000) {
            lastReport = now;
            log(`downloading ${Math.round((received / total) * 100)}%`);
          }
        });
        response.pipe(file);
        file.on('finish', () => file.close(() => resolve(destination)));
        file.on('error', reject);
      })
      .on('error', reject);
  });
}

// Windows-only package, so PowerShell is always available and this stays
// dependency-free. Node has no built-in zip reader.
function extract(archive, destination) {
  const result = spawnSync(
    'powershell.exe',
    [
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-Command',
      `Expand-Archive -LiteralPath '${archive.replace(/'/g, "''")}' -DestinationPath '${destination.replace(/'/g, "''")}' -Force`,
    ],
    { stdio: ['ignore', 'ignore', 'pipe'] }
  );

  if (result.status !== 0) {
    const detail = (result.stderr || '').toString().trim();
    throw new Error(`Failed to extract ${archive}${detail ? `: ${detail}` : ''}`);
  }
}

function assertLooksLikeZip(file) {
  const size = fs.statSync(file).size;
  if (size < 1024 * 1024) {
    throw new Error(
      `Downloaded archive is only ${size} bytes, which cannot be a self-contained build. ` +
        'The release asset is probably missing.'
    );
  }
  const header = Buffer.alloc(2);
  const fd = fs.openSync(file, 'r');
  try {
    fs.readSync(fd, header, 0, 2, 0);
  } finally {
    fs.closeSync(fd);
  }
  if (header[0] !== 0x50 || header[1] !== 0x4b) {
    throw new Error('Downloaded file is not a ZIP archive. The release asset URL may be wrong.');
  }
}

async function ensureRuntime() {
  if (process.platform !== 'win32') {
    throw new Error(
      'mcp-server-ppt only runs on Windows: it drives PowerPoint through the native COM API.'
    );
  }

  if (process.env.MCP_SERVER_PPT_HOME) {
    const local = path.join(process.env.MCP_SERVER_PPT_HOME, EXE_NAME);
    if (!fs.existsSync(local)) {
      throw new Error(`MCP_SERVER_PPT_HOME is set but ${local} does not exist.`);
    }
    return local;
  }

  const version = packageVersion();
  const root = installRoot(version);
  const exe = path.join(root, EXE_NAME);

  if (fs.existsSync(exe)) {
    return exe;
  }

  const url = assetUrl(version);
  log(`installing the PowerPoint MCP server ${version} (about 64 MB to download, 145 MB on disk, one time)`);
  log(`source: ${url}`);

  fs.mkdirSync(root, { recursive: true });
  const staging = fs.mkdtempSync(path.join(os.tmpdir(), 'mcp-server-ppt-'));
  const archive = path.join(staging, assetName(version));

  try {
    await download(url, archive);
    assertLooksLikeZip(archive);
    extract(archive, root);

    if (!fs.existsSync(exe)) {
      throw new Error(`Archive extracted but ${EXE_NAME} was not found in it.`);
    }
    log('install complete');
    return exe;
  } catch (error) {
    // Never leave a half-extracted directory behind: the next run would find no
    // exe, re-download, and Expand-Archive -Force would merge the two.
    fs.rmSync(root, { recursive: true, force: true });
    throw error;
  } finally {
    fs.rmSync(staging, { recursive: true, force: true });
  }
}

module.exports = {
  EXE_NAME,
  assetName,
  assetUrl,
  ensureRuntime,
  installRoot,
  packageVersion,
};
