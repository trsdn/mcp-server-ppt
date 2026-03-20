import {
  existsSync,
  mkdirSync,
  openSync,
  closeSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from "fs";
import { execFileSync } from "child_process";
import { basename, dirname, extname, join, resolve } from "path";
import {
  DEFAULT_EXECUTE_TIMEOUT_MS,
  DEFAULT_MODEL,
  DEFAULT_PLAN_TIMEOUT_MS,
  DEFAULT_VERIFY_TIMEOUT_MS,
  REPO_ROOT,
} from "./constants.mjs";
import { parsePlanFromText } from "./planner.mjs";
import { createRuntime, destroyRuntime, runPhase } from "./runtime.mjs";
import { extractTextRunsFromSlideXml, findMissingRequiredTexts, findSlideQualityIssues } from "./validation.mjs";

const MAX_REPAIR_ATTEMPTS = 3;

function buildBusinessQualityRules() {
  return [
    "- Treat business-slide quality as a hard requirement, not a nice-to-have.",
    "- Do NOT use novelty or decorative shapes like sun, star, heart, cloud, moon, or smiley shapes unless the plan explicitly asks for them.",
    "- For KPI cards, panels, and callouts, prefer rectangles or rounded rectangles with flat fills and simple lines.",
    "- Use a restrained business palette: neutral background, dark text, one main accent, and semantic red/green only where the content explicitly calls for risk or positive next steps.",
    "- If the result looks like default PowerPoint theme art or contains gaudy styling, replace it before finishing.",
  ];
}

function withoutExtension(filePath) {
  return basename(filePath, extname(filePath));
}

function defaultOutputPath() {
  const now = new Date();
  const stamp = [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, "0"),
    String(now.getDate()).padStart(2, "0"),
    "-",
    String(now.getHours()).padStart(2, "0"),
    String(now.getMinutes()).padStart(2, "0"),
    String(now.getSeconds()).padStart(2, "0"),
  ].join("");

  return resolve(process.cwd(), `ppt-mcp-agent-${stamp}.pptx`);
}

function preparePaths(outputPath, overwrite) {
  const resolvedOutputPath = resolve(outputPath || defaultOutputPath());
  const planPath = join(dirname(resolvedOutputPath), `${withoutExtension(resolvedOutputPath)}.plan.json`);
  const artifactsDir = join(dirname(resolvedOutputPath), `${withoutExtension(resolvedOutputPath)}-artifacts`);
  const summaryPath = join(artifactsDir, "run-summary.json");

  const conflictingPath = [resolvedOutputPath, planPath, artifactsDir].find((path) => existsSync(path));
  if (conflictingPath && !overwrite) {
    throw new Error(`Output artifacts already exist: ${conflictingPath}. Use --overwrite to replace them.`);
  }

  if (overwrite) {
    rmSync(resolvedOutputPath, { force: true });
    rmSync(planPath, { force: true });
    rmSync(artifactsDir, { recursive: true, force: true });
  }

  mkdirSync(dirname(resolvedOutputPath), { recursive: true });
  mkdirSync(artifactsDir, { recursive: true });

  return {
    outputPath: resolvedOutputPath,
    planPath,
    artifactsDir,
    summaryPath,
  };
}

function loadArchetypeIds() {
  const archetypesDir = join(REPO_ROOT, "src", "PptMcp.Core", "Data", "archetypes");

  return readdirSync(archetypesDir)
    .filter((fileName) => fileName.endsWith(".md"))
    .map((fileName) => fileName.replace(/\.md$/i, ""))
    .filter((id) => id !== "registry" && id !== "evidence-design")
    .sort();
}

function readText(relativePath) {
  return readFileSync(join(REPO_ROOT, relativePath), "utf-8").trim();
}

function loadPlanningGuidance() {
  const registry = readText("src\\PptMcp.Core\\Data\\archetypes\\registry.md");
  const generationPipeline = readText("skills\\shared\\generation-pipeline.md");

  return [
    "Use this repository guidance while planning:",
    "",
    "## Archetype Registry",
    registry,
    "",
    "## Generation Pipeline",
    generationPipeline,
  ].join("\n");
}

function buildPlanningPrompt({ task, archetypeIds }) {
  return [
    "You are the planning phase of a PowerPoint deck agent.",
    "Do not create or modify any presentation in this phase.",
    "Do not rely on MCP batch execution or subagents.",
    "Return ONLY valid JSON and nothing else.",
    "",
    "Required schema:",
    "{",
    '  "slides": [',
    "    {",
    '      "index": 1,',
    '      "title": "Action title",',
    '      "archetypeId": "executive-summary",',
    '      "intent": "What the slide must help the audience understand or decide",',
    '      "content": "Detailed build instructions specific enough for an execution phase"',
    "    }",
    "  ]",
    "}",
    "",
    `Allowed archetypeIds: ${archetypeIds.join(", ")}`,
    "",
    loadPlanningGuidance(),
    "",
    "User task:",
    task,
  ].join("\n");
}

function loadPlanFromFile(filePath) {
  const resolvedPath = resolve(filePath);
  if (!existsSync(resolvedPath)) {
    throw new Error(`Plan file was not found: ${resolvedPath}`);
  }

  const content = readFileSync(resolvedPath, "utf-8");
  const plan = parsePlanFromText(content);
  if (!plan) {
    throw new Error(`Plan file did not contain a valid deck plan: ${resolvedPath}`);
  }

  return plan;
}

function buildSlideExecutionRules(plan) {
  const blankFriendlyArchetypes = new Set([
    "appendix",
    "big-number",
    "chart-insight-callout",
    "column-bar-chart",
    "comparison",
    "executive-summary",
    "framework",
    "kpi-card-dashboard",
    "operational-kpi",
    "process-diagram",
    "recommendations",
    "simple-table",
    "timeline-roadmap",
    "waterfall-chart",
  ]);

  return plan.slides.flatMap((slide) => {
    const rules = [`- Slide ${slide.index}: archetype '${slide.archetypeId}'.`];
    const content = slide.content || "";
    const contentLower = content.toLowerCase();
    const minimumShapeMatch = content.match(/(\d+)\+\s*shapes/i);

    rules.push(`- Slide ${slide.index}: render the exact slide title text "${slide.title}" as a visible heading and preserve it through later edits.`);

    if (slide.archetypeId === "title-slide") {
      rules.push(`- Slide ${slide.index}: use slide(action='create', layout_name='Title Slide').`);
      rules.push(`- Slide ${slide.index}: prefer placeholders for title and subtitle.`);
      return rules;
    }

    rules.push(`- Slide ${slide.index}: do NOT use slide(action='create', layout_name='Title Slide').`);

    if (contentLower.includes("blank layout") || blankFriendlyArchetypes.has(slide.archetypeId)) {
      rules.push(`- Slide ${slide.index}: create the slide with slide(action='create', layout_name='Blank').`);
    }

    rules.push(`- Slide ${slide.index}: implement the detailed content literally; do not collapse it into only a title and subtitle.`);
    rules.push(`- Slide ${slide.index}: use separate shapes/text boxes/containers for distinct panels, cards, and callouts.`);

    if (minimumShapeMatch) {
      rules.push(`- Slide ${slide.index}: do not finish below ${minimumShapeMatch[1]} shapes because the plan explicitly requires that density.`);
    }

    if (contentLower.includes("kpi card") || contentLower.includes("kpi cards")) {
      rules.push(`- Slide ${slide.index}: build distinct KPI cards, each with its own background shape and text elements.`);
    }

    if (contentLower.includes("clustered column chart") || contentLower.includes("chart")) {
      rules.push(`- Slide ${slide.index}: create a real chart object when chart data is specified, not a text placeholder describing a chart.`);
    }

    if (contentLower.includes("insight panel")) {
      rules.push(`- Slide ${slide.index}: build the insight panel as its own container with separate bullet text elements.`);
    }

    if (contentLower.includes("callout")) {
      rules.push(`- Slide ${slide.index}: build each callout as a separate colored box with its own text.`);
    }

    return rules;
  });
}

function buildExecutionPrompt({ task, plan, outputPath, showPowerPoint }) {
  const slideExecutionRules = buildSlideExecutionRules(plan);
  const businessQualityRules = buildBusinessQualityRules();

  return [
    "You are the execution phase of a PowerPoint deck agent.",
    "You are operating through mcp-server-ppt only.",
    "Do not rely on MCP batch execution or subagents.",
    "Treat the plan as fixed input. Build slide-by-slide with normal sequential MCP tool calls.",
    "",
    "Execution rules:",
    `- Create a new presentation at this exact path: ${outputPath}`,
    `- When creating the file, set show=${showPowerPoint ? "true" : "false"}`,
    "- Keep one PowerPoint session open for the full build.",
    "- Use the skill guidance plus design tools as needed.",
    "- Build slides in plan order.",
    `- The final presentation MUST contain exactly ${plan.slides.length} slide(s).`,
    "- Prefer targeted edits over delete/rebuild.",
    "- Before finishing, verify slide count with slide list/read operations.",
    "- Finish only after file(action='close', save=true).",
    "",
    "Required MCP tool pattern:",
    "- Start with file(action='create', path=..., show=...)",
    "- For each slide, create the slide first, then populate content",
    "- Use slide(action='list') to confirm the deck structure before closing",
    "- Prefer placeholder(action='set-text') when a layout already provides title/subtitle placeholders",
    "",
    "Business design quality rules:",
    ...businessQualityRules,
    "",
    "Slide-specific execution rules:",
    ...slideExecutionRules,
    "",
    "Title-slide recipe:",
    "- If the archetype is title-slide, create slide(action='create', layout_name='Title Slide')",
    "- Then use placeholder(action='list')",
    "- Set the title/subtitle with placeholder(action='set-text', placeholder_index=..., text=...)",
    "- Only fall back to freeform text boxes if the title layout does not expose placeholders",
    "",
    "Original user task:",
    task,
    "",
    "Structured plan:",
    "```json",
    JSON.stringify(plan, null, 2),
    "```",
    "",
    "Return a concise summary of what was built and any unresolved risks.",
  ].join("\n");
}

function buildRepairPrompt({ task, plan, outputPath, validationError, showPowerPoint }) {
  const slideExecutionRules = buildSlideExecutionRules(plan);
  const businessQualityRules = buildBusinessQualityRules();

  return [
    "You are the repair phase of a PowerPoint deck agent.",
    "A previous execution produced an incomplete presentation.",
    "Repair the presentation through mcp-server-ppt only.",
    "Do not rely on MCP batch execution or subagents.",
    "",
    "Repair goal:",
    `- Output file path: ${outputPath}`,
    `- Required final slide count: ${plan.slides.length}`,
    `- Validation failure to fix: ${validationError}`,
    "- If the validation failure names missing required text elements, add those text elements literally and preserve everything already built correctly.",
    "- If the validation failure names quality issues such as novelty shapes or palette problems, restyle the slide until those issues are gone.",
    "- If the file exists, open and repair it. If it is missing, create it.",
    `- When creating a new file, set show=${showPowerPoint ? "true" : "false"}`,
    "- Build or repair slides so the final deck matches the plan.",
    "- Do not stop while a planned callout or footer container exists without its required text.",
    "- Use slide list/read operations before closing to confirm the final structure.",
    "- Finish only after file(action='close', save=true).",
    "",
    "Business design quality rules:",
    ...businessQualityRules,
    "",
    "Slide-specific repair rules:",
    ...slideExecutionRules,
    "",
    "Repair recipe for title-slide outputs:",
    "- If the file is empty or missing slides, open or create it",
    "- Use slide(action='create', layout_name='Title Slide') for title-slide plan items",
    "- Use placeholder(action='list') and placeholder(action='set-text') to write title and subtitle",
    "- Confirm the final slide count with slide(action='list') before closing",
    "- If a non-title slide is missing its planned heading, add the exact slide title from the plan as a visible top heading before closing",
    "",
    "Original user task:",
    task,
    "",
    "Structured plan:",
    "```json",
    JSON.stringify(plan, null, 2),
    "```",
    "",
    "Return a concise summary of what was repaired.",
  ].join("\n");
}

function buildVerificationPrompt({ task, outputPath, artifactsDir, plan }) {
  const businessQualityRules = buildBusinessQualityRules();
  const slideExecutionRules = buildSlideExecutionRules(plan);

  return [
    "You are the verification phase of a PowerPoint deck agent.",
    "Re-open the generated presentation and review it with normal sequential MCP tool calls.",
    "Do not rely on MCP batch execution or subagents.",
    "Apply only targeted fixes for obvious structural issues.",
    "Preserve all content that already matches the plan, especially planned headings, callouts, and footer text.",
    "",
    "Verification rules:",
    `- Open this presentation: ${outputPath}`,
    "- Inspect slides with slide list/read plus shape/text inspection as needed.",
    `- Export slide images for human review into this directory: ${artifactsDir}`,
    "- Focus on both structure and visual business quality.",
    "- Specifically look for novelty shapes, default PowerPoint theme art, weak palette choices, poor alignment, unreadable density, and obviously unprofessional styling.",
    "- If you find a visual quality issue, fix it instead of merely reporting it.",
    "- Before finishing, confirm that each slide still contains the exact planned title text and any other required literal text from the plan.",
    "",
    "Business design quality rules:",
    ...businessQualityRules,
    "",
    "Slide-specific verification rules:",
    ...slideExecutionRules,
    "",
    "Structured plan:",
    "```json",
    JSON.stringify(plan, null, 2),
    "```",
    "- Save and close the presentation when done.",
    "",
    "Original user task:",
    task,
    "",
    "Return a concise verification report with:",
    "- what was checked",
    "- what was fixed",
    "- any remaining concerns",
  ].join("\n");
}

function writeJson(path, value) {
  writeFileSync(path, `${JSON.stringify(value, null, 2)}\n`, "utf-8");
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function escapePowerShellSingleQuoted(text) {
  return String(text).replace(/'/g, "''");
}

function runPowerShellScript(script) {
  return execFileSync(
    "powershell.exe",
    ["-NoProfile", "-NonInteractive", "-Command", script],
    {
      encoding: "utf-8",
      stdio: ["ignore", "pipe", "pipe"],
      windowsHide: true,
    }
  ).trim();
}

function closePresentationIfOpen(filePath) {
  const target = escapePowerShellSingleQuoted(filePath);
  const script = [
    "Add-Type -AssemblyName Microsoft.VisualBasic",
    `$target = '${target}'`,
    "$targetItem = Get-Item $target -ErrorAction SilentlyContinue",
    "if ($null -eq $targetItem) { exit 0 }",
    "$targetPath = $targetItem.FullName",
    "try { $app = [Microsoft.VisualBasic.Interaction]::GetObject('', 'PowerPoint.Application') } catch { exit 0 }",
    "try {",
    "  for ($i = $app.Presentations.Count; $i -ge 1; $i--) {",
    "    $presentation = $app.Presentations.Item($i)",
    "    try {",
    "      $presentationPath = $null",
    "      try {",
    "        if ($presentation.FullName) {",
    "          $presentationPath = (Get-Item $presentation.FullName -ErrorAction SilentlyContinue).FullName",
    "        }",
    "      } catch {}",
    "      if ($presentationPath -eq $targetPath) {",
    "        try { $presentation.Save() } catch {}",
    "        $presentation.Close()",
    "      }",
    "    } finally {",
    "      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)",
    "    }",
    "  }",
    "} finally {",
    "  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)",
    "}",
  ].join("; ");

  try {
    runPowerShellScript(script);
  } catch {
    // Best-effort cleanup only.
  }
}

async function waitForExpectedSlideCount(filePath, expectedSlideCount, timeoutMs = 15000) {
  const startedAt = Date.now();
  let lastCount = 0;

  while ((Date.now() - startedAt) < timeoutMs) {
    try {
      lastCount = getPptxSlideCount(filePath);
      if (lastCount >= expectedSlideCount) {
        return lastCount;
      }
    } catch {
      // Keep retrying until PowerPoint has flushed the file.
    }

    await delay(500);
  }

  return lastCount;
}

async function waitForFileUnlock(filePath, timeoutMs = 15000) {
  const startedAt = Date.now();

  while ((Date.now() - startedAt) < timeoutMs) {
    try {
      const handle = openSync(filePath, "r+");
      closeSync(handle);
      return true;
    } catch {
      await delay(500);
    }
  }

  return false;
}

function getPptxSlideCount(filePath) {
  const target = escapePowerShellSingleQuoted(filePath);
  const script = [
    "Add-Type -AssemblyName System.IO.Compression.FileSystem",
    `$target = '${target}'`,
    "$zip = [System.IO.Compression.ZipFile]::OpenRead($target)",
    "try {",
    "  ($zip.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide\\d+\\.xml$' }).Count",
    "} finally {",
    "  $zip.Dispose()",
    "}",
  ].join("; ");

  const output = runPowerShellScript(script);
  const count = Number.parseInt(output, 10);
  if (!Number.isFinite(count)) {
    throw new Error(`Could not determine slide count for ${filePath}.`);
  }

  return count;
}

function getSlideXmlFromPptx(filePath, slideIndex) {
  const target = escapePowerShellSingleQuoted(filePath);
  const entryName = escapePowerShellSingleQuoted(`ppt/slides/slide${slideIndex}.xml`);
  const script = [
    "Add-Type -AssemblyName System.IO.Compression.FileSystem",
    `$target = '${target}'`,
    `$entryName = '${entryName}'`,
    "$zip = [System.IO.Compression.ZipFile]::OpenRead($target)",
    "try {",
    "  $entry = $zip.GetEntry($entryName)",
    "  if ($null -eq $entry) { throw \"Missing slide entry: $entryName\" }",
    "  $reader = New-Object System.IO.StreamReader($entry.Open())",
    "  try {",
    "    $content = $reader.ReadToEnd()",
    "    [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($content))",
    "  } finally {",
    "    $reader.Dispose()",
    "  }",
    "} finally {",
    "  $zip.Dispose()",
    "}",
  ].join("; ");

  const base64 = runPowerShellScript(script);
  return Buffer.from(base64, "base64").toString("utf-8");
}

async function validateDeckArtifact(outputPath, expectedSlideCount) {
  if (!existsSync(outputPath)) {
    throw new Error(`Output file was not created: ${outputPath}`);
  }

  closePresentationIfOpen(outputPath);

  const unlocked = await waitForFileUnlock(outputPath);
  if (!unlocked) {
    throw new Error(`Output file is still locked after execution: ${outputPath}`);
  }

  const slideCount = await waitForExpectedSlideCount(outputPath, expectedSlideCount);
  if (slideCount < expectedSlideCount) {
    throw new Error(
      `Output file contains ${slideCount} slide(s), but ${expectedSlideCount} were expected.`
    );
  }

  return { slideCount };
}

async function validateDeckContent(outputPath, plan) {
  const problems = [];

  for (const slide of plan.slides) {
    const slideXml = getSlideXmlFromPptx(outputPath, slide.index);
    const actualTexts = extractTextRunsFromSlideXml(slideXml);
    const missingTexts = findMissingRequiredTexts(slide, actualTexts);
    const qualityIssues = findSlideQualityIssues(slide, slideXml);

    if (missingTexts.length > 0) {
      problems.push(
        `Slide ${slide.index} is missing required text elements: ${missingTexts.map((text) => `"${text}"`).join(", ")}`
      );
    }

    problems.push(...qualityIssues);
  }

  if (problems.length > 0) {
    throw new Error(problems.join(" "));
  }
}

async function validateDeckOutput(outputPath, plan) {
  await validateDeckArtifact(outputPath, plan.slides.length);
  await validateDeckContent(outputPath, plan);
}

export async function runDeckAgent(options) {
  const model = options.model || DEFAULT_MODEL;
  const paths = preparePaths(options.outputPath, Boolean(options.overwrite));
  const archetypeIds = loadArchetypeIds();
  const task = options.task || "Execute the supplied deck plan exactly.";
  let plan;

  if (options.planFilePath) {
    plan = loadPlanFromFile(options.planFilePath);
  } else {
    const planRuntime = await createRuntime({
      model,
      verbose: options.verbose,
      enableMcp: false,
    });

    let planResult;
    try {
      planResult = await runPhase(planRuntime, {
        name: "plan",
        label: "Planning",
        enableMcp: false,
        enableSkills: true,
        timeoutMs: options.planTimeoutMs || DEFAULT_PLAN_TIMEOUT_MS,
        prompt: buildPlanningPrompt({
          task,
          archetypeIds,
        }),
      });
    } finally {
      await destroyRuntime(planRuntime);
    }

    if (!planResult.ok) {
      throw new Error(`Planning failed: ${planResult.error}`);
    }

    plan = parsePlanFromText(planResult.content);
    if (!plan) {
      throw new Error("Planning phase did not return a valid deck plan.");
    }
  }

  writeJson(paths.planPath, plan);

  const executeRuntime = await createRuntime({
    model,
    verbose: options.verbose,
    mcpServerPath: options.mcpServerPath,
  });

  let executeResult;
  let repairSummary = null;
  try {
    executeResult = await runPhase(executeRuntime, {
      name: "execute",
      label: "Execution",
      enableMcp: true,
      enableSkills: true,
      successArtifactPath: paths.outputPath,
      timeoutMs: options.executeTimeoutMs || DEFAULT_EXECUTE_TIMEOUT_MS,
      prompt: buildExecutionPrompt({
        task,
        plan,
        outputPath: paths.outputPath,
        showPowerPoint: Boolean(options.showPowerPoint),
      }),
    });
  } finally {
    await destroyRuntime(executeRuntime);
  }

  if (!executeResult.ok) {
    if (!(executeResult.isTimeout && executeResult.artifactDetected)) {
      throw new Error(`Execution failed: ${executeResult.error}`);
    }
  }

  let deckValidationError = null;
  try {
    await validateDeckOutput(paths.outputPath, plan);
  } catch (validationError) {
    deckValidationError = validationError;
  }

  if (deckValidationError) {
    const repairSummaries = [];

    for (let attempt = 1; attempt <= MAX_REPAIR_ATTEMPTS; attempt++) {
      const repairRuntime = await createRuntime({
        model,
        verbose: options.verbose,
        mcpServerPath: options.mcpServerPath,
      });

      try {
        const repairResult = await runPhase(repairRuntime, {
          name: "improve",
          label: `Repair ${attempt}/${MAX_REPAIR_ATTEMPTS}`,
          enableMcp: true,
          enableSkills: true,
          successArtifactPath: paths.outputPath,
          timeoutMs: options.executeTimeoutMs || DEFAULT_EXECUTE_TIMEOUT_MS,
          prompt: buildRepairPrompt({
            task,
            plan,
            outputPath: paths.outputPath,
            validationError: deckValidationError instanceof Error ? deckValidationError.message : String(deckValidationError),
            showPowerPoint: Boolean(options.showPowerPoint),
          }),
        });

        if (!repairResult.ok) {
          if (!(repairResult.isTimeout && repairResult.artifactDetected)) {
            throw new Error(`Repair failed: ${repairResult.error}`);
          }
        }

        repairSummaries.push(repairResult.content || repairResult.error || `Repair attempt ${attempt} completed.`);
      } finally {
        await destroyRuntime(repairRuntime);
      }

      try {
        await validateDeckOutput(paths.outputPath, plan);
        repairSummary = repairSummaries.join("\n\n");
        deckValidationError = null;
        break;
      } catch (validationError) {
        deckValidationError = validationError;
      }
    }

    if (deckValidationError) {
      throw deckValidationError;
    }
  }

  let verifyResult = null;
  if (!options.skipVerify) {
    const verifyRuntime = await createRuntime({
      model,
      verbose: options.verbose,
      mcpServerPath: options.mcpServerPath,
    });

    try {
      verifyResult = await runPhase(verifyRuntime, {
        name: "verify",
        label: "Verification",
        enableMcp: true,
        enableSkills: true,
        successArtifactPath: paths.outputPath,
        timeoutMs: options.verifyTimeoutMs || DEFAULT_VERIFY_TIMEOUT_MS,
        prompt: buildVerificationPrompt({
          task,
          outputPath: paths.outputPath,
          artifactsDir: paths.artifactsDir,
          plan,
        }),
      });
    } finally {
      await destroyRuntime(verifyRuntime);
    }

    if (!verifyResult.ok) {
      if (!(verifyResult.isTimeout && verifyResult.artifactDetected)) {
        throw new Error(`Verification failed: ${verifyResult.error}`);
      }
    }

    if (!existsSync(paths.outputPath)) {
      throw new Error(`Verification completed but output file is missing: ${paths.outputPath}`);
    }

    await validateDeckOutput(paths.outputPath, plan);
  }

  const summary = {
    task: options.task,
    model,
    outputPath: paths.outputPath,
    planPath: paths.planPath,
    artifactsDir: paths.artifactsDir,
    plan,
    executionSummary: executeResult.content || executeResult.error,
    repairSummary,
    verificationSummary: verifyResult?.content || null,
    constraints: {
      noMcpBatchDependency: true,
      noSubagentDependency: true,
    },
  };

  writeJson(paths.summaryPath, summary);

  return summary;
}
