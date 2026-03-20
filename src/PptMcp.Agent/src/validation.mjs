const BUSINESS_STYLE_ARCHETYPES = new Set([
  "appendix",
  "big-number",
  "chart-insight-callout",
  "column-bar-chart",
  "comparison",
  "executive-summary",
  "framework",
  "kpi-card-dashboard",
  "operational-kpi",
  "recommendations",
  "simple-table",
  "timeline-roadmap",
  "waterfall-chart",
]);

const NOVELTY_PRESET_SHAPES = new Set([
  "arc",
  "bevel",
  "chevron",
  "cloud",
  "decagon",
  "donut",
  "gear6",
  "gear9",
  "heart",
  "hexagon",
  "moon",
  "smileyFace",
  "star10",
  "star12",
  "star16",
  "star24",
  "star32",
  "star4",
  "star5",
  "star6",
  "star7",
  "star8",
  "sun",
]);

const MAX_VIVID_FILL_COLORS = 3;

function normalizeText(text) {
  return String(text || "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function decodeXmlEntities(text) {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function collectQuotedTexts(text) {
  const values = [];

  for (const match of text.matchAll(/"([^"]+)"/g)) {
    values.push(match[1]);
  }

  return values;
}

export function collectRequiredSlideTexts(slide) {
  const requiredTexts = new Set();

  if (slide?.title) {
    requiredTexts.add(slide.title);
  }

  const content = slide?.content || "";
  const bulletAnchorIndex = content.indexOf("Use these bullets:");
  if (bulletAnchorIndex >= 0) {
    for (const value of collectQuotedTexts(content.slice(bulletAnchorIndex))) {
      requiredTexts.add(value);
    }
  } else {
    for (const match of content.matchAll(/(?:sentence|with|footer[^:]*):\s*"([^"]+)"/gi)) {
      requiredTexts.add(match[1]);
    }
  }

  return [...requiredTexts];
}

export function extractTextRunsFromSlideXml(xml) {
  const values = [];

  for (const match of xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)) {
    const value = decodeXmlEntities(match[1]).trim();
    if (value) {
      values.push(value);
    }
  }

  return values;
}

export function extractPresetGeometryNamesFromSlideXml(xml) {
  const geometries = [];

  for (const match of xml.matchAll(/<a:prstGeom[^>]*prst="([^"]+)"/g)) {
    geometries.push(match[1]);
  }

  return geometries;
}

export function extractSolidFillColorsFromSlideXml(xml) {
  const colors = [];

  for (const match of xml.matchAll(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/g)) {
    const srgbMatch = match[1].match(/<a:srgbClr[^>]*val="([0-9A-Fa-f]{6})"/);
    if (srgbMatch) {
      colors.push(srgbMatch[1].toUpperCase());
    }
  }

  return colors;
}

function isBusinessStyleSlide(slide) {
  return BUSINESS_STYLE_ARCHETYPES.has(slide?.archetypeId || "");
}

function isVividHexColor(hex) {
  const red = Number.parseInt(hex.slice(0, 2), 16);
  const green = Number.parseInt(hex.slice(2, 4), 16);
  const blue = Number.parseInt(hex.slice(4, 6), 16);
  const max = Math.max(red, green, blue);
  const min = Math.min(red, green, blue);

  if (!Number.isFinite(max) || !Number.isFinite(min) || max === 0) {
    return false;
  }

  const saturation = (max - min) / max;
  return saturation >= 0.35 && (max - min) >= 40 && max >= 80;
}

function getHueFamily(hex) {
  const red = Number.parseInt(hex.slice(0, 2), 16) / 255;
  const green = Number.parseInt(hex.slice(2, 4), 16) / 255;
  const blue = Number.parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(red, green, blue);
  const min = Math.min(red, green, blue);
  const delta = max - min;

  if (delta === 0) {
    return "neutral";
  }

  let hue;
  if (max === red) {
    hue = ((green - blue) / delta) % 6;
  } else if (max === green) {
    hue = ((blue - red) / delta) + 2;
  } else {
    hue = ((red - green) / delta) + 4;
  }

  const degrees = ((hue * 60) + 360) % 360;
  if (degrees < 30 || degrees >= 330) {
    return "red";
  }
  if (degrees < 75) {
    return "orange";
  }
  if (degrees < 150) {
    return "green";
  }
  if (degrees < 210) {
    return "cyan";
  }
  if (degrees < 270) {
    return "blue";
  }

  return "purple";
}

export function findSlideQualityIssues(slide, slideXml) {
  if (!isBusinessStyleSlide(slide)) {
    return [];
  }

  const issues = [];
  const noveltyShapes = [...new Set(
    extractPresetGeometryNamesFromSlideXml(slideXml).filter((shape) => NOVELTY_PRESET_SHAPES.has(shape))
  )];

  if (noveltyShapes.length > 0) {
    issues.push(
      `Slide ${slide.index} uses novelty preset shapes that are not acceptable for a business slide: ${noveltyShapes.join(", ")}. Replace them with simple rectangles or rounded rectangles.`
    );
  }

  const vividFillColors = [...new Set(
    extractSolidFillColorsFromSlideXml(slideXml).filter(isVividHexColor)
  )];
  const vividColorFamilies = [...new Set(vividFillColors.map(getHueFamily).filter((family) => family !== "neutral"))];

  if (vividColorFamilies.length > MAX_VIVID_FILL_COLORS) {
    issues.push(
      `Slide ${slide.index} uses too many distinct vivid color families for a business slide: ${vividFillColors.join(", ")}. Use a restrained palette with neutrals plus one main accent and semantic red/green only where justified.`
    );
  }

  return issues;
}

export function findMissingRequiredTexts(slide, actualTexts) {
  const combinedActualText = normalizeText(actualTexts.join(" "));

  return collectRequiredSlideTexts(slide).filter((requiredText) => {
    return !combinedActualText.includes(normalizeText(requiredText));
  });
}
