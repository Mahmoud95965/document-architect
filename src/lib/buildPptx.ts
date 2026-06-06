import pptxgen from "pptxgenjs";
import type {
  FreestylePresentationJSON,
  FreestyleSlide,
  SlideElement,
  HeroTextElement,
  SubtextElement,
  GlassCardElement,
  ImageNodeElement,
  KpiBlockElement,
  AccentShapeElement,
  DividerLineElement,
} from "./generate.functions";
import { isProjectArabic } from "./lang";

// ─────────────────────────────────────────────────────────────
//  EXPANDED IMAGE LIBRARY  (topic-keyed Unsplash fallbacks)
// ─────────────────────────────────────────────────────────────
const IMAGE_MAP: Record<string, string> = {
  // Business / Finance
  corporate:    "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?w=1400&auto=format&fit=crop&q=85",
  business:     "https://images.unsplash.com/photo-1664575602554-2087b04935a5?w=1400&auto=format&fit=crop&q=85",
  finance:      "https://images.unsplash.com/photo-1611974789855-9c2a0a7236a3?w=1400&auto=format&fit=crop&q=85",
  investment:   "https://images.unsplash.com/photo-1559526324-593bc073d938?w=1400&auto=format&fit=crop&q=85",
  economy:      "https://images.unsplash.com/photo-1611974789855-9c2a0a7236a3?w=1400&auto=format&fit=crop&q=85",
  market:       "https://images.unsplash.com/photo-1611974789855-9c2a0a7236a3?w=1400&auto=format&fit=crop&q=85",
  revenue:      "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=1400&auto=format&fit=crop&q=85",
  budget:       "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=1400&auto=format&fit=crop&q=85",
  growth:       "https://images.unsplash.com/photo-1543286386-2e659306cd6c?w=1400&auto=format&fit=crop&q=85",
  profit:       "https://images.unsplash.com/photo-1543286386-2e659306cd6c?w=1400&auto=format&fit=crop&q=85",
  // Technology
  tech:         "https://images.unsplash.com/photo-1518770660439-4636190af475?w=1400&auto=format&fit=crop&q=85",
  technology:   "https://images.unsplash.com/photo-1518770660439-4636190af475?w=1400&auto=format&fit=crop&q=85",
  digital:      "https://images.unsplash.com/photo-1526374965328-7f61d4dc18c5?w=1400&auto=format&fit=crop&q=85",
  ai:           "https://images.unsplash.com/photo-1677442135703-1787eea5ce01?w=1400&auto=format&fit=crop&q=85",
  software:     "https://images.unsplash.com/photo-1587620962725-abab19836bd2?w=1400&auto=format&fit=crop&q=85",
  code:         "https://images.unsplash.com/photo-1587620962725-abab19836bd2?w=1400&auto=format&fit=crop&q=85",
  data:         "https://images.unsplash.com/photo-1551288049-bebda4e38f71?w=1400&auto=format&fit=crop&q=85",
  cloud:        "https://images.unsplash.com/photo-1544197150-b99a580bb7a8?w=1400&auto=format&fit=crop&q=85",
  cyber:        "https://images.unsplash.com/photo-1550751827-4bd374c3f58b?w=1400&auto=format&fit=crop&q=85",
  // People / Teams
  team:         "https://images.unsplash.com/photo-1522071820081-009f0129c71c?w=1400&auto=format&fit=crop&q=85",
  people:       "https://images.unsplash.com/photo-1529156069898-49953e39b3ac?w=1400&auto=format&fit=crop&q=85",
  leadership:   "https://images.unsplash.com/photo-1507003211169-0a1dd7228f2d?w=1400&auto=format&fit=crop&q=85",
  management:   "https://images.unsplash.com/photo-1576267423048-15c0040fec78?w=1400&auto=format&fit=crop&q=85",
  // Real Estate
  real:         "https://images.unsplash.com/photo-1560518883-ce09059eeffa?w=1400&auto=format&fit=crop&q=85",
  estate:       "https://images.unsplash.com/photo-1560518883-ce09059eeffa?w=1400&auto=format&fit=crop&q=85",
  property:     "https://images.unsplash.com/photo-1560518883-ce09059eeffa?w=1400&auto=format&fit=crop&q=85",
  building:     "https://images.unsplash.com/photo-1486325212027-8081e485255e?w=1400&auto=format&fit=crop&q=85",
  architecture: "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?w=1400&auto=format&fit=crop&q=85",
  city:         "https://images.unsplash.com/photo-1477959858617-67f85cf4f1df?w=1400&auto=format&fit=crop&q=85",
  // Healthcare
  health:       "https://images.unsplash.com/photo-1576091160550-2173dba999ef?w=1400&auto=format&fit=crop&q=85",
  medical:      "https://images.unsplash.com/photo-1576091160550-2173dba999ef?w=1400&auto=format&fit=crop&q=85",
  hospital:     "https://images.unsplash.com/photo-1519494026892-80bbd2d6fd0d?w=1400&auto=format&fit=crop&q=85",
  // Education
  education:    "https://images.unsplash.com/photo-1523050854058-8df90110c9f1?w=1400&auto=format&fit=crop&q=85",
  learning:     "https://images.unsplash.com/photo-1509062522246-3755977927d7?w=1400&auto=format&fit=crop&q=85",
  university:   "https://images.unsplash.com/photo-1541339907198-e08756dedf3f?w=1400&auto=format&fit=crop&q=85",
  // Nature / Environment
  nature:       "https://images.unsplash.com/photo-1441974231531-c6227db76b6e?w=1400&auto=format&fit=crop&q=85",
  environment:  "https://images.unsplash.com/photo-1500534314209-a25ddb2bd429?w=1400&auto=format&fit=crop&q=85",
  green:        "https://images.unsplash.com/photo-1473448912268-2022ce9509d8?w=1400&auto=format&fit=crop&q=85",
  energy:       "https://images.unsplash.com/photo-1473341304170-971dccb5ac1e?w=1400&auto=format&fit=crop&q=85",
  // Marketing / Creative
  marketing:    "https://images.unsplash.com/photo-1432888498266-38ffec3eaf0a?w=1400&auto=format&fit=crop&q=85",
  social:       "https://images.unsplash.com/photo-1611162617474-5b21e879e113?w=1400&auto=format&fit=crop&q=85",
  creative:     "https://images.unsplash.com/photo-1558618666-fcd25c85cd64?w=1400&auto=format&fit=crop&q=85",
  design:       "https://images.unsplash.com/photo-1558618666-fcd25c85cd64?w=1400&auto=format&fit=crop&q=85",
  strategy:     "https://images.unsplash.com/photo-1460925895917-afdab827c52f?w=1400&auto=format&fit=crop&q=85",
  // Travel / Logistics
  travel:       "https://images.unsplash.com/photo-1436491865332-7a61a109cc05?w=1400&auto=format&fit=crop&q=85",
  logistics:    "https://images.unsplash.com/photo-1586528116311-ad8dd3c8310d?w=1400&auto=format&fit=crop&q=85",
  supply:       "https://images.unsplash.com/photo-1586528116311-ad8dd3c8310d?w=1400&auto=format&fit=crop&q=85",
  // Default
  default:      "https://images.unsplash.com/photo-1557804506-669a67965ba0?w=1400&auto=format&fit=crop&q=85",
  cover:        "https://images.unsplash.com/photo-1557804506-669a67965ba0?w=1400&auto=format&fit=crop&q=85",
};

export function getUnsplashImageUrl(query = ""): string {
  const q = query.toLowerCase();
  for (const [key, url] of Object.entries(IMAGE_MAP)) {
    if (q.includes(key)) return url;
  }
  return IMAGE_MAP.default;
}

// ─────────────────────────────────────────────────────────────
//  COLOUR HELPERS
// ─────────────────────────────────────────────────────────────
function hexToRgb(hex: string): { r: number; g: number; b: number } {
  const h = hex.replace("#", "");
  const n = parseInt(h, 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}

function darkened(hex: string, factor = 0.4): string {
  const { r, g, b } = hexToRgb(hex);
  const d = (v: number) => Math.max(0, Math.floor(v * factor)).toString(16).padStart(2, "0");
  return `${d(r)}${d(g)}${d(b)}`;
}

/** Strips leading # and returns 6-char hex */
function toHex(color: string | undefined, fallback: string): string {
  if (!color) return fallback.replace("#", "");
  return color.replace("#", "");
}

/** Resolves a background image URL — either the AI-provided direct URL or a fallback */
function resolveBgUrl(slide: FreestyleSlide): string {
  if (slide.bgImagePath && slide.bgImagePath.startsWith("http")) {
    return slide.bgImagePath;
  }
  // Fallback: derive from slide_title keywords
  return getUnsplashImageUrl(slide.slide_title);
}

// ─────────────────────────────────────────────────────────────
//  ELEMENT RENDERER — the core of the Generative Canvas Engine
// ─────────────────────────────────────────────────────────────
function renderElement(
  pptx: pptxgen,
  slide: pptxgen.Slide,
  el: SlideElement,
  isArabic: boolean,
  fontFace: string,
  accentHex: string,
): void {
  const align = isArabic ? "right" : "left";
  const rtlMode = isArabic;

  // 1. حسابات أبعاد آمنة لمنع خروج العناصر برة الشاشة
  const safeX = Math.min(Math.max(el.x, 0.5), 12.0);
  const safeY = Math.min(Math.max(el.y, 0.5), 6.5);
  const safeW = Math.min(el.w, 13.0 - safeX);
  const safeH = Math.min(el.h, 7.0 - safeY);

  switch (el.type) {

    // ── Hero Text ─────────────────────────────────────────────
    case "hero_text": {
      const e = el as HeroTextElement;
      slide.addText(e.text, {
        x: safeX, y: safeY, w: safeW, h: safeH,
        fontSize: e.fontSize && e.fontSize > 40 ? 38 : (e.fontSize || 38),
        bold: true,
        color: toHex(e.color, "#FFFFFF"),
        fontFace, align, rtlMode,
        fit: "shrink",
        shadow: { type: "outer", color: "000000", blur: 10, offset: 4, angle: 45, opacity: 0.6 },
      });
      break;
    }

    // ── Subtext ───────────────────────────────────────────────
    case "subtext": {
      const e = el as SubtextElement;
      slide.addText(e.text, {
        x: safeX, y: safeY, w: safeW, h: safeH,
        fontSize: e.fontSize ?? 14,
        bold: false,
        color: toHex(e.color, "#94A3B8"),
        fontFace, align, rtlMode,
        fit: "shrink",
      });
      break;
    }

    // ── Glass Card ────────────────────────────────────────────
    case "glass_card": {
      const e = el as GlassCardElement;

      // طبقة التعتيم (Dark Scrim) خلف الكارت مباشرة
      slide.addShape(pptx.ShapeType.roundRect, {
        x: safeX, y: safeY, w: safeW, h: safeH,
        fill: { color: "000000", transparency: 60 },
        line: { color: accentHex, width: 1.5, transparency: 30 },
        rectRadius: 0.2,
      });

      // Accent top bar
      slide.addShape(pptx.ShapeType.roundRect, {
        x: safeX, y: safeY, w: safeW, h: 0.07,
        fill: { color: accentHex },
        rectRadius: 0.05,
      });

      // صب النصوص داخل الكارت مع حماية الحجم والمحاذاة
      slide.addText([
        ...(e.title ? [{ text: `${e.title}\n`, options: { fontSize: 16, bold: true, color: accentHex, fontFace } }] : []),
        { text: e.text || "", options: { fontSize: 12, color: "CBD5E1", fontFace } }
      ], {
        x: safeX + 0.2,
        y: safeY + 0.2,
        w: safeW - 0.4,
        h: safeH - 0.4,
        align, rtlMode,
        fit: "shrink",
        valign: "top",
      });
      break;
    }

    // ── Image Node ────────────────────────────────────────────
    case "image_node": {
      const e = el as ImageNodeElement;
      const imgUrl = e.imageUrl?.startsWith("http") ? e.imageUrl : getUnsplashImageUrl(e.imageUrl);

      slide.addShape(pptx.ShapeType.roundRect, {
        x: safeX - 0.05, y: safeY - 0.05, w: safeW + 0.1, h: safeH + 0.1,
        fill: { color: accentHex, transparency: 85 },
        rectRadius: 0.18,
      });

      slide.addImage({
        path: imgUrl,
        x: safeX, y: safeY, w: safeW, h: safeH,
        sizing: { type: "cover", w: safeW, h: safeH },
        rounding: true,
      });
      break;
    }

    // ── KPI Block ─────────────────────────────────────────────
    case "kpi_block": {
      const e = el as KpiBlockElement;

      slide.addShape(pptx.ShapeType.roundRect, {
        x: safeX, y: safeY, w: safeW, h: safeH,
        fill: { color: darkened(accentHex, 0.18), transparency: 20 },
        line: { color: accentHex, width: 0.5, transparency: 55 },
        rectRadius: 0.18,
      });

      slide.addShape(pptx.ShapeType.roundRect, {
        x: safeX, y: safeY, w: 0.06, h: safeH,
        fill: { color: accentHex },
        rectRadius: 0.05,
      });

      slide.addText([
        { text: `${e.value || ""}\n`, options: { fontSize: 44, bold: true, color: accentHex, fontFace } },
        { text: e.label || "", options: { fontSize: 14, color: "94A3B8", fontFace } }
      ], {
        x: safeX + 0.2, y: safeY + 0.1, w: safeW - 0.3, h: safeH - 0.2,
        align, rtlMode,
        fit: "shrink",
        valign: "middle",
      });
      break;
    }

    // ── Accent Shape ──────────────────────────────────────────
    case "accent_shape": {
      const e = el as AccentShapeElement;
      const shapeColor = toHex(e.color, `#${accentHex}`);
      const trans = Math.round((1 - (e.opacity ?? 0.15)) * 100);
      const shapeType: pptxgen.SHAPE_NAME = e.shape === "rect" ? "rect" : 'ellipse';

      slide.addShape(shapeType, {
        x: safeX, y: safeY, w: safeW, h: safeH,
        fill: { color: shapeColor, transparency: trans },
      });
      break;
    }

    // ── Divider Line ──────────────────────────────────────────
    case "divider_line": {
      const e = el as DividerLineElement;
      slide.addShape(pptx.ShapeType.rect, {
        x: safeX, y: safeY, w: safeW, h: Math.max(safeH, 0.025),
        fill: { color: toHex(e.color, "#334155") },
      });
      break;
    }
  }
}

// ─────────────────────────────────────────────────────────────
//  PREMIUM COVER SLIDE
// ─────────────────────────────────────────────────────────────
function buildCoverSlide(
  pptx: pptxgen,
  data: FreestylePresentationJSON,
  isArabic: boolean,
  fontFace: string,
) {
  const slide = pptx.addSlide();
  const align = isArabic ? "right" : "left";
  const rtlMode = isArabic;

  // Cover image — derive from first slide's bgImagePath or slide_title
  const firstSlide = data.slides[0];
  const coverImg = firstSlide?.bgImagePath?.startsWith("http")
    ? firstSlide.bgImagePath
    : getUnsplashImageUrl(firstSlide?.slide_title ?? data.title);

  // Background — dark premium
  slide.background = { fill: "060D1B" };

  // Full-bleed cover image with heavy overlay
  slide.addImage({ path: coverImg, x: 0, y: 0, w: 13.33, h: 7.5, sizing: { type: "cover", w: 13.33, h: 7.5 } });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 7.5, fill: { color: "000000", transparency: 15 } });

  // Left gradient panel overlay
  const panelSide = isArabic ? 7.5 : 0;
  slide.addShape(pptx.ShapeType.rect, { x: panelSide, y: 0, w: 5.83, h: 7.5, fill: { color: "060D1B", transparency: 5 } });

  // Giant accent geometric decoration
  slide.addShape('ellipse', {
    x: isArabic ? -1.5 : 11.0, y: -1.5, w: 5.5, h: 5.5,
    fill: { color: "6366F1", transparency: 80 },
  });
  slide.addShape('ellipse', {
    x: isArabic ? 0.0 : 12.5, y: 3.0, w: 2.8, h: 2.8,
    fill: { color: "38BDF8", transparency: 70 },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: isArabic ? 13.15 : 0, y: 0, w: 0.18, h: 7.5,
    fill: { color: "6366F1" },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: isArabic ? 5.7 : 5.55, y: 0, w: 0.05, h: 7.5,
    fill: { color: "FFFFFF", transparency: 85 },
  });

  // Category badge
  slide.addShape(pptx.ShapeType.roundRect, {
    x: isArabic ? 7.9 : 0.55, y: 1.3, w: 2.1, h: 0.45,
    fill: { color: "6366F1" }, rectRadius: 0.14,
  });
  slide.addText("TOLZY FLOW", {
    x: isArabic ? 7.9 : 0.55, y: 1.3, w: 2.1, h: 0.45,
    fontSize: 9, bold: true, color: "FFFFFF", fontFace, align: "center", charSpacing: 2,
  });

  // Title — grand and bold
  slide.addText(data.title, {
    x: isArabic ? 7.8 : 0.55, y: 2.05, w: 5.0, h: 2.6,
    fontSize: 38, bold: true, color: "FFFFFF", fontFace, align, rtlMode,
    lineSpacing: 44,
    shadow: { type: "outer", color: "000000", blur: 10, offset: 4, angle: 45, opacity: 0.7 },
  });

  // Subtitle
  if (data.subtitle) {
    slide.addText(data.subtitle, {
      x: isArabic ? 7.8 : 0.55, y: 4.85, w: 5.0, h: 0.8,
      fontSize: 15, color: "94A3B8", fontFace, align, rtlMode, lineSpacing: 20,
    });
  }

  // Thin accent line below title
  slide.addShape(pptx.ShapeType.rect, {
    x: isArabic ? 7.9 : 0.55, y: 4.72, w: 2.0, h: 0.06,
    fill: { color: "6366F1" },
  });

  // Slide count badge bottom right
  slide.addShape(pptx.ShapeType.roundRect, {
    x: isArabic ? 0.5 : 11.6, y: 6.75, w: 1.18, h: 0.42,
    fill: { color: "FFFFFF", transparency: 15 }, rectRadius: 0.12,
  });
  slide.addText(`${data.slides.length} Slides`, {
    x: isArabic ? 0.5 : 11.6, y: 6.75, w: 1.18, h: 0.42,
    fontSize: 9, color: "FFFFFF", fontFace, align: "center", bold: true,
  });

  // Branding label
  slide.addText("AI-Generated Presentation", {
    x: isArabic ? 7.8 : 0.55, y: 5.75, w: 5.0, h: 0.38,
    fontSize: 10, color: "475569", fontFace, align, rtlMode,
  });
}

// ─────────────────────────────────────────────────────────────
//  CONTENT SLIDE BUILDER
// ─────────────────────────────────────────────────────────────
function buildContentSlide(
  pptx: pptxgen,
  slideData: FreestyleSlide,
  isArabic: boolean,
  fontFace: string,
  slideIndex: number,
): void {
  const slide = pptx.addSlide();
  const accentHex = toHex(slideData.accentColor, "#6366F1");
  const bgUrl = resolveBgUrl(slideData);

  // ── Background ────────────────────────────────────────────
  // Always render the bg image first, full-bleed
  slide.addImage({
    path: bgUrl,
    x: 0, y: 0, w: 13.33, h: 7.5,
    sizing: { type: "cover", w: 13.33, h: 7.5 },
  });

  // Deep dark overlay for readability — dynamically adjusted
  // Slides with heavy image use stronger overlay
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 13.33, h: 7.5,
    fill: { color: "000000", transparency: 30 },
  });

  // Subtle gradient vignette on bottom edge
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 4.5, w: 13.33, h: 3.0,
    fill: { color: "000000", transparency: 45 },
  });

  // Accent neon line — top of slide
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 13.33, h: 0.055,
    fill: { color: accentHex },
  });

  // Slide number badge — bottom left
  slide.addShape(pptx.ShapeType.roundRect, {
    x: isArabic ? 12.0 : 0.6, y: 7.0, w: 0.72, h: 0.32,
    fill: { color: accentHex, transparency: 20 }, rectRadius: 0.08,
  });
  slide.addText(`${String(slideIndex + 1).padStart(2, "0")}`, {
    x: isArabic ? 12.0 : 0.6, y: 7.0, w: 0.72, h: 0.32,
    fontSize: 9, bold: true, color: "FFFFFF", fontFace, align: "center",
  });

  // TOLZY branding
  slide.addText("TOLZY Flow", {
    x: isArabic ? 0.5 : 11.4, y: 7.1, w: 1.4, h: 0.25,
    fontSize: 7, color: "FFFFFF", fontFace,
    align: isArabic ? "left" : "right",
  });

  // ── Render all AI-defined elements ────────────────────────
  for (const el of (slideData.elements ?? [])) {
    try {
      renderElement(pptx, slide, el, isArabic, fontFace, accentHex);
    } catch (err) {
      // Skip malformed elements gracefully — don't crash the whole deck
      console.warn(`[buildPptx] Skipped element type="${el.type}" at (${el.x},${el.y}):`, err);
    }
  }
}

// ─────────────────────────────────────────────────────────────
//  MAIN EXPORT
// ─────────────────────────────────────────────────────────────
export async function generateAndDownloadPptx(
  data: FreestylePresentationJSON,
  fileName = "presentation.pptx",
) {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";

  const isArabic = isProjectArabic(data as any);
  const fontFace = "Calibri";

  // ── Cover ──
  buildCoverSlide(pptx, data, isArabic, fontFace);

  // ── Content slides ──
  data.slides.forEach((slideData, index) => {
    buildContentSlide(pptx, slideData, isArabic, fontFace, index);
  });

  await pptx.writeFile({ fileName });
}




