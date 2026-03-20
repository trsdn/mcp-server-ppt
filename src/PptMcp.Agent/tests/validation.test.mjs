import assert from "node:assert/strict";
import test from "node:test";

import {
  collectRequiredSlideTexts,
  extractPresetGeometryNamesFromSlideXml,
  extractSolidFillColorsFromSlideXml,
  extractTextRunsFromSlideXml,
  findSlideQualityIssues,
  findMissingRequiredTexts,
} from "../src/validation.mjs";

test("collectRequiredSlideTexts includes title, bullets, callouts, and footer from a dashboard plan", () => {
  const slide = {
    index: 1,
    title: "Approve a West recovery plan to protect otherwise strong regional growth",
    content: [
      "Use these bullets: \"North and East contribute most of the Q4 acceleration.\"; \"West remains below Q1 revenue and is about $0.8M behind plan.\".",
      "Below the insight panel, add one red risk callout box with a bold lead-in and one sentence: \"Risk: If West stays flat next quarter, the annual target gap widens by about $1.6M.\".",
      "Beneath that, add one green next-step callout box with: \"Next step: Approve a 90-day West recovery plan and weekly executive review cadence.\".",
      "Add a small 8pt gray footer at the bottom: \"Source: FY2026 regional revenue tracker; dashboard as of Q4 close.\"",
    ].join(" "),
  };

  assert.deepEqual(collectRequiredSlideTexts(slide), [
    "Approve a West recovery plan to protect otherwise strong regional growth",
    "North and East contribute most of the Q4 acceleration.",
    "West remains below Q1 revenue and is about $0.8M behind plan.",
    "Risk: If West stays flat next quarter, the annual target gap widens by about $1.6M.",
    "Next step: Approve a 90-day West recovery plan and weekly executive review cadence.",
    "Source: FY2026 regional revenue tracker; dashboard as of Q4 close.",
  ]);
});

test("extractTextRunsFromSlideXml returns decoded text runs", () => {
  const xml = `
    <p:txBody>
      <a:p><a:r><a:t>Risk:</a:t></a:r></a:p>
      <a:p><a:r><a:t>If West stays flat &amp; delivery slips &lt;2 pts&gt;</a:t></a:r></a:p>
    </p:txBody>
  `;

  assert.deepEqual(extractTextRunsFromSlideXml(xml), [
    "Risk:",
    "If West stays flat & delivery slips <2 pts>",
  ]);
});

test("findMissingRequiredTexts matches combined slide text across split text boxes", () => {
  const slide = {
    index: 1,
    title: "Approve a West recovery plan to protect otherwise strong regional growth",
    content: [
      "Use these bullets: \"North and East contribute most of the Q4 acceleration.\".",
      "Below the insight panel, add one red risk callout box with a bold lead-in and one sentence: \"Risk: If West stays flat next quarter, the annual target gap widens by about $1.6M.\".",
      "Beneath that, add one green next-step callout box with: \"Next step: Approve a 90-day West recovery plan and weekly executive review cadence.\".",
      "Add a small 8pt gray footer at the bottom: \"Source: FY2026 regional revenue tracker; dashboard as of Q4 close.\"",
    ].join(" "),
  };

  const actualTexts = [
    "Approve a West recovery plan to protect otherwise strong regional growth",
    "North and East contribute most of the Q4 acceleration.",
    "Risk:",
    "If West stays flat next quarter, the annual target gap widens by about $1.6M.",
    "Next step: Approve a 90-day West recovery plan and weekly executive review cadence.",
  ];

  assert.deepEqual(findMissingRequiredTexts(slide, actualTexts), [
    "Source: FY2026 regional revenue tracker; dashboard as of Q4 close.",
  ]);
});

test("extractPresetGeometryNamesFromSlideXml returns preset shape names", () => {
  const xml = `
    <p:sp><p:spPr><a:prstGeom prst="sun"><a:avLst/></a:prstGeom></p:spPr></p:sp>
    <p:sp><p:spPr><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></p:spPr></p:sp>
  `;

  assert.deepEqual(extractPresetGeometryNamesFromSlideXml(xml), ["sun", "roundRect"]);
});

test("extractSolidFillColorsFromSlideXml returns explicit solid fill colors", () => {
  const xml = `
    <p:sp><p:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr></p:sp>
    <p:sp><p:spPr><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></p:spPr></p:sp>
  `;

  assert.deepEqual(extractSolidFillColorsFromSlideXml(xml), ["4472C4", "C00000"]);
});

test("findSlideQualityIssues flags novelty shapes on business slides", () => {
  const slide = {
    index: 1,
    archetypeId: "kpi-card-dashboard",
  };

  const xml = `
    <p:sp><p:spPr><a:prstGeom prst="sun"><a:avLst/></a:prstGeom></p:spPr></p:sp>
    <p:sp><p:spPr><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></p:spPr></p:sp>
  `;

  assert.deepEqual(findSlideQualityIssues(slide, xml), [
    "Slide 1 uses novelty preset shapes that are not acceptable for a business slide: sun. Replace them with simple rectangles or rounded rectangles.",
  ]);
});

test("findSlideQualityIssues flags overly colorful business slides", () => {
  const slide = {
    index: 1,
    archetypeId: "executive-summary",
  };

  const xml = `
    <p:sp><p:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr></p:sp>
    <p:sp><p:spPr><a:solidFill><a:srgbClr val="00B0F0"/></a:solidFill></p:spPr></p:sp>
    <p:sp><p:spPr><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></p:spPr></p:sp>
    <p:sp><p:spPr><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></p:spPr></p:sp>
  `;

  assert.deepEqual(findSlideQualityIssues(slide, xml), [
    "Slide 1 uses too many distinct vivid color families for a business slide: 4472C4, 00B0F0, 70AD47, C00000. Use a restrained palette with neutrals plus one main accent and semantic red/green only where justified.",
  ]);
});

test("findSlideQualityIssues allows semantic red and green shades without flagging palette sprawl", () => {
  const slide = {
    index: 1,
    archetypeId: "kpi-card-dashboard",
  };

  const xml = `
    <p:sp><p:spPr><a:solidFill><a:srgbClr val="86EFAC"/></a:solidFill></p:spPr></p:sp>
    <p:sp><p:spPr><a:solidFill><a:srgbClr val="166534"/></a:solidFill></p:spPr></p:sp>
    <p:sp><p:spPr><a:solidFill><a:srgbClr val="B91C1C"/></a:solidFill></p:spPr></p:sp>
    <p:sp><p:spPr><a:solidFill><a:srgbClr val="991B1B"/></a:solidFill></p:spPr></p:sp>
  `;

  assert.deepEqual(findSlideQualityIssues(slide, xml), []);
});
