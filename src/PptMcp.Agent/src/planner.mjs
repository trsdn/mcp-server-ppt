function normalizeSlidePlan(value, index) {
  if (!value || typeof value !== "object") {
    return null;
  }

  const candidate = value;
  const normalized = {
    index: Number.isFinite(candidate.index) ? Number(candidate.index) : index + 1,
    title: String(candidate.title || "").trim(),
    archetypeId: String(candidate.archetypeId || candidate.layout || "").trim(),
    intent: String(candidate.intent || "").trim(),
    content: String(candidate.content || candidate.notes || "").trim(),
  };

  if (!normalized.title || !normalized.archetypeId || !normalized.content) {
    return null;
  }

  return normalized;
}

function normalizePlan(slides) {
  const normalizedSlides = slides
    .map((slide, index) => normalizeSlidePlan(slide, index))
    .filter(Boolean)
    .sort((left, right) => left.index - right.index)
    .map((slide, index) => ({ ...slide, index: index + 1 }));

  if (normalizedSlides.length === 0) {
    return null;
  }

  return { slides: normalizedSlides };
}

function coerceToPlan(value) {
  if (value && typeof value === "object") {
    if (Array.isArray(value.slides)) {
      return normalizePlan(value.slides);
    }

    if (value.plan && typeof value.plan === "object" && Array.isArray(value.plan.slides)) {
      return normalizePlan(value.plan.slides);
    }
  }

  if (Array.isArray(value)) {
    return normalizePlan(value);
  }

  return null;
}

function extractOutermostJson(text) {
  const startIndex = text.search(/[{[]/);
  if (startIndex === -1) {
    return null;
  }

  const openChar = text[startIndex];
  const closeChar = openChar === "{" ? "}" : "]";
  let depth = 0;
  let inString = false;
  let escape = false;

  for (let index = startIndex; index < text.length; index++) {
    const char = text[index];

    if (escape) {
      escape = false;
      continue;
    }

    if (char === "\\" && inString) {
      escape = true;
      continue;
    }

    if (char === "\"") {
      inString = !inString;
      continue;
    }

    if (inString) {
      continue;
    }

    if (char === openChar) {
      depth++;
    }

    if (char === closeChar) {
      depth--;
    }

    if (depth === 0) {
      return text.slice(startIndex, index + 1);
    }
  }

  return null;
}

function parseMarkdownPlan(text) {
  const slides = [];

  const blockPattern =
    /###\s*Slide\s+(\d+)[:\s]*([^\n]+)\n(?:[\s\S]*?-\s*\*{0,2}Archetype\*{0,2}[:\s]+([^\n]+)\n)?(?:[\s\S]*?-\s*\*{0,2}Intent\*{0,2}[:\s]+([^\n]+)\n)?(?:[\s\S]*?-\s*\*{0,2}Content\*{0,2}[:\s]+([^\n]+))?/gi;

  let match;
  while ((match = blockPattern.exec(text)) !== null) {
    slides.push({
      index: Number(match[1]),
      title: match[2].trim().replace(/\*{1,2}/g, ""),
      archetypeId: (match[3] || "").trim().replace(/\*{1,2}/g, ""),
      intent: (match[4] || "").trim().replace(/\*{1,2}/g, ""),
      content: (match[5] || match[4] || match[2] || "").trim().replace(/\*{1,2}/g, ""),
    });
  }

  if (slides.length > 0) {
    return normalizePlan(slides);
  }

  return null;
}

export function parsePlanFromText(text) {
  const candidates = [];

  const codeBlockRegex = /```(?:json)?\s*([\s\S]*?)```/g;
  let codeBlockMatch;
  while ((codeBlockMatch = codeBlockRegex.exec(text)) !== null) {
    const inner = codeBlockMatch[1].trim();
    if (inner.startsWith("{") || inner.startsWith("[")) {
      candidates.push(inner);
    }
  }

  const outerJson = extractOutermostJson(text);
  if (outerJson) {
    candidates.push(outerJson);
  }

  candidates.push(text.trim());

  for (const candidate of candidates) {
    try {
      const parsed = JSON.parse(candidate);
      const plan = coerceToPlan(parsed);
      if (plan) {
        return plan;
      }
    } catch {
      // Ignore parse failures and continue to the next extraction strategy.
    }
  }

  return parseMarkdownPlan(text);
}
