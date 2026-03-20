#!/usr/bin/env node

import { DEFAULT_EXECUTE_TIMEOUT_MS, DEFAULT_MODEL, DEFAULT_PLAN_TIMEOUT_MS, DEFAULT_VERIFY_TIMEOUT_MS } from "./constants.mjs";

function printHelp() {
  console.log(`ppt-mcp-agent

Official source-side Copilot SDK client for orchestrating mcp-server-ppt.

Usage:
  ppt-mcp-agent run --task "Build a 5-slide deck on ..." [options]

Options:
  --task <text>                 Natural-language deck request
  --plan-file <path>            Reuse a precomputed plan file and skip planning
  --output <path>               Output PPTX path (default: timestamped file in cwd)
  --model <name>                Copilot model to use (default: ${DEFAULT_MODEL})
  --mcp-server <path>           Override MCP server executable path
  --show                        Show PowerPoint while executing
  --overwrite                   Replace existing output/artifact files
  --skip-verify                 Skip the verification phase
  --verbose                     Print tool-call event traces
  --plan-timeout-ms <number>    Planning timeout in ms (default: ${DEFAULT_PLAN_TIMEOUT_MS})
  --execute-timeout-ms <number> Execution timeout in ms (default: ${DEFAULT_EXECUTE_TIMEOUT_MS})
  --verify-timeout-ms <number>  Verification timeout in ms (default: ${DEFAULT_VERIFY_TIMEOUT_MS})
  -h, --help                    Show help
`);
}

function parseInteger(value, flagName) {
  const parsed = Number.parseInt(value, 10);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error(`${flagName} must be a positive integer.`);
  }

  return parsed;
}

function readNextValue(argv, index, flagName) {
  const value = argv[index + 1];
  if (!value || value.startsWith("-")) {
    throw new Error(`${flagName} requires a value.`);
  }

  return value;
}

function parseArgs(argv) {
  const args = {
    command: null,
    task: null,
    planFilePath: null,
    outputPath: null,
    model: DEFAULT_MODEL,
    mcpServerPath: null,
    showPowerPoint: false,
    overwrite: false,
    skipVerify: false,
    verbose: false,
    planTimeoutMs: DEFAULT_PLAN_TIMEOUT_MS,
    executeTimeoutMs: DEFAULT_EXECUTE_TIMEOUT_MS,
    verifyTimeoutMs: DEFAULT_VERIFY_TIMEOUT_MS,
  };

  for (let index = 0; index < argv.length; index++) {
    const arg = argv[index];

    if (!args.command && !arg.startsWith("-")) {
      args.command = arg;
      continue;
    }

    switch (arg) {
      case "--task":
        args.task = readNextValue(argv, index, "--task");
        index++;
        break;
      case "--output":
        args.outputPath = readNextValue(argv, index, "--output");
        index++;
        break;
      case "--plan-file":
        args.planFilePath = readNextValue(argv, index, "--plan-file");
        index++;
        break;
      case "--model":
        args.model = readNextValue(argv, index, "--model");
        index++;
        break;
      case "--mcp-server":
        args.mcpServerPath = readNextValue(argv, index, "--mcp-server");
        index++;
        break;
      case "--show":
        args.showPowerPoint = true;
        break;
      case "--overwrite":
        args.overwrite = true;
        break;
      case "--skip-verify":
        args.skipVerify = true;
        break;
      case "--verbose":
        args.verbose = true;
        break;
      case "--plan-timeout-ms":
        args.planTimeoutMs = parseInteger(readNextValue(argv, index, "--plan-timeout-ms"), "--plan-timeout-ms");
        index++;
        break;
      case "--execute-timeout-ms":
        args.executeTimeoutMs = parseInteger(readNextValue(argv, index, "--execute-timeout-ms"), "--execute-timeout-ms");
        index++;
        break;
      case "--verify-timeout-ms":
        args.verifyTimeoutMs = parseInteger(readNextValue(argv, index, "--verify-timeout-ms"), "--verify-timeout-ms");
        index++;
        break;
      case "-h":
      case "--help":
        args.help = true;
        break;
      default:
        throw new Error(`Unknown argument: ${arg}`);
    }
  }

  return args;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));

  if (args.help || !args.command) {
    printHelp();
    return;
  }

  if (args.command !== "run") {
    throw new Error(`Unknown command: ${args.command}`);
  }

  if (!args.task && !args.planFilePath) {
    throw new Error("Either --task or --plan-file is required for the run command.");
  }

  let runDeckAgent;
  try {
    ({ runDeckAgent } = await import("./orchestrator.mjs"));
  } catch (error) {
    if (error && error.code === "ERR_MODULE_NOT_FOUND") {
      throw new Error(
        "Missing runtime dependency. Run 'npm install' in src\\PptMcp.Agent before using the agent."
      );
    }

    throw error;
  }

  const summary = await runDeckAgent(args);

  console.log("\n=== Agent Summary ===");
  console.log(`Output: ${summary.outputPath}`);
  console.log(`Plan: ${summary.planPath}`);
  console.log(`Artifacts: ${summary.artifactsDir}`);
  console.log(`Slides planned: ${summary.plan.slides.length}`);
}

main().catch((error) => {
  console.error(`ppt-mcp-agent failed: ${error instanceof Error ? error.message : String(error)}`);
  process.exit(1);
});
