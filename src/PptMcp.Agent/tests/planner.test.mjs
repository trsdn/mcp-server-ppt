import assert from "node:assert/strict";
import test from "node:test";

import { parsePlanFromText } from "../src/planner.mjs";

test("parsePlanFromText parses a direct JSON plan", () => {
  const plan = parsePlanFromText(`{
    "slides": [
      {
        "index": 1,
        "title": "Executive summary",
        "archetypeId": "executive-summary",
        "intent": "Summarize the decision",
        "content": "Three bullets covering outcome, risk, and ask"
      }
    ]
  }`);

  assert.deepEqual(plan, {
    slides: [
      {
        index: 1,
        title: "Executive summary",
        archetypeId: "executive-summary",
        intent: "Summarize the decision",
        content: "Three bullets covering outcome, risk, and ask",
      },
    ],
  });
});

test("parsePlanFromText unwraps a nested plan object and renumbers slides", () => {
  const plan = parsePlanFromText(`{
    "plan": {
      "slides": [
        {
          "index": 3,
          "title": "Next actions",
          "archetypeId": "recommendations",
          "intent": "Show owners and timing",
          "content": "A three-row action table"
        },
        {
          "index": 1,
          "title": "Title slide",
          "archetypeId": "title-slide",
          "intent": "Introduce the story",
          "content": "Title plus subtitle"
        }
      ]
    }
  }`);

  assert.deepEqual(plan?.slides.map((slide) => ({ index: slide.index, title: slide.title })), [
    { index: 1, title: "Title slide" },
    { index: 2, title: "Next actions" },
  ]);
});

test("parsePlanFromText extracts JSON from a fenced code block", () => {
  const plan = parsePlanFromText(`
Here is the plan:

\`\`\`json
[
  {
    "title": "Market context",
    "layout": "comparison",
    "intent": "Frame the change drivers",
    "notes": "Two-column comparison of before and after conditions"
  }
]
\`\`\`
`);

  assert.equal(plan?.slides[0].archetypeId, "comparison");
  assert.equal(plan?.slides[0].content, "Two-column comparison of before and after conditions");
});

test("parsePlanFromText falls back to markdown slide blocks", () => {
  const plan = parsePlanFromText(`
### Slide 1: Recommendation
- Archetype: recommendations
- Intent: Get approval for the next phase
- Content: Three action tiles with owner and timeline
`);

  assert.deepEqual(plan, {
    slides: [
      {
        index: 1,
        title: "Recommendation",
        archetypeId: "recommendations",
        intent: "Get approval for the next phase",
        content: "Three action tiles with owner and timeline",
      },
    ],
  });
});

test("parsePlanFromText returns null for incomplete plans", () => {
  const plan = parsePlanFromText(`{"slides":[{"title":"Only title"}]}`);
  assert.equal(plan, null);
});
