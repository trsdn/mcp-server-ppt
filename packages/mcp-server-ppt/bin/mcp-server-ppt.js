#!/usr/bin/env node
'use strict';

// stdout belongs to the MCP transport. Diagnostics go to stderr, always.

const { spawn } = require('child_process');
const { ensureRuntime, packageVersion } = require('../lib/runtime');

async function main() {
  const args = process.argv.slice(2);

  if (args[0] === '--install') {
    const exe = await ensureRuntime();
    process.stderr.write(`[mcp-server-ppt] ready: ${exe}\n`);
    return 0;
  }

  if (args[0] === '--version') {
    process.stderr.write(`mcp-server-ppt ${packageVersion()}\n`);
    return 0;
  }

  const exe = await ensureRuntime();

  const child = spawn(exe, args, { stdio: 'inherit', windowsHide: true });

  // Forward termination so a client killing the wrapper does not strand a
  // PowerPoint-owning server process.
  for (const signal of ['SIGINT', 'SIGTERM', 'SIGHUP']) {
    process.on(signal, () => {
      if (!child.killed) {
        child.kill();
      }
    });
  }

  return new Promise((resolve, reject) => {
    child.on('error', reject);
    child.on('exit', (code, signal) => resolve(signal ? 1 : code === null ? 1 : code));
  });
}

main().then(
  (code) => {
    process.exitCode = code;
  },
  (error) => {
    process.stderr.write(`[mcp-server-ppt] ${error.message}\n`);
    process.exitCode = 1;
  }
);
