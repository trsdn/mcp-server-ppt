'use strict';

// Network-free checks for the wrapper. Run with `npm test` in this directory.
//
// The check that matters is the last one: the asset name this wrapper requests
// has to match the asset name the release workflow produces. Those two strings
// live in different files, in different languages, and nothing else connects
// them, so a rename on either side would only ever surface as a 404 on a user's
// machine after publish.

const assert = require('assert');
const fs = require('fs');
const path = require('path');

const runtime = require('./runtime');
const pkg = require('../package.json');

const failures = [];

function check(name, fn) {
  try {
    fn();
    process.stdout.write(`  ok   ${name}\n`);
  } catch (error) {
    failures.push(name);
    process.stdout.write(`  FAIL ${name}\n         ${error.message}\n`);
  }
}

check('package declares Windows-only so npm refuses other platforms', () => {
  assert.deepStrictEqual(pkg.os, ['win32']);
  assert.deepStrictEqual(pkg.cpu, ['x64']);
});

check('every file listed in "files" exists', () => {
  for (const entry of pkg.files) {
    const target = path.join(__dirname, '..', entry);
    if (entry === 'README.md' && !fs.existsSync(target)) {
      throw new Error('README.md is missing');
    }
    if (entry !== 'README.md') {
      assert.ok(fs.existsSync(target), `${entry} is missing`);
    }
  }
});

check('bin entry point exists and is executable JavaScript', () => {
  const bin = path.join(__dirname, '..', pkg.bin['mcp-server-ppt']);
  assert.ok(fs.existsSync(bin), `${bin} is missing`);
  assert.ok(fs.readFileSync(bin, 'utf8').startsWith('#!/usr/bin/env node'));
});

check('asset URL points at this package version in this repository', () => {
  const url = runtime.assetUrl();
  assert.ok(url.startsWith('https://github.com/trsdn/mcp-server-ppt/releases/download/v'), url);
  assert.ok(url.includes(pkg.version), url);
});

check('MCP_SERVER_PPT_ASSET_URL overrides the default source', () => {
  process.env.MCP_SERVER_PPT_ASSET_URL = 'https://example.invalid/x.zip';
  try {
    assert.strictEqual(runtime.assetUrl(), 'https://example.invalid/x.zip');
  } finally {
    delete process.env.MCP_SERVER_PPT_ASSET_URL;
  }
});

check('asset name matches the archive the release workflow builds', () => {
  const workflow = path.join(__dirname, '..', '..', '..', '.github', 'workflows', 'release.yml');
  if (!fs.existsSync(workflow)) {
    // Running from an installed package rather than the repository.
    return;
  }
  const contents = fs.readFileSync(workflow, 'utf8');
  const expected = runtime.assetName('$version');
  assert.ok(
    contents.includes(expected),
    `release.yml does not produce "${expected}" - the wrapper would download a 404`
  );
});

process.stdout.write(failures.length ? `\n${failures.length} check(s) failed\n` : '\nall checks passed\n');
process.exitCode = failures.length ? 1 : 0;
