import { createServerFn } from "@tanstack/react-start";

export type DocSection = {
  heading: string;
  subheading?: string;
  paragraphs?: string[];
  table?: { headers: string[]; rows: string[][] };
};

export type DocJSON = {
  title: string;
  subtitle?: string;
  sections: DocSection[];
};

export type ExcelCell = {
  value: string | number;
  formula?: string; // e.g. "=SUM(C2:C5)"
  bold?: boolean;
  bg?: string; // Hex color code without #
  color?: string; // Hex color code without #
};

export type ExcelSheet = {
  name: string;
  headers: string[];
  rows: ExcelCell[][];
};

export type ExcelJSON = {
  title: string;
  sheets: ExcelSheet[];
};

// ─── Legacy (kept for Word/Excel unchanged) ───────────────────
export type SlideItem = {
  title: string;
  bullets: string[];
  bgGradient?: { from: string; to: string };
  imageQuery?: string;
  layoutHint?: number;
};

// ─── Generative Canvas Schema ─────────────────────────────────
export type SlideElementType =
  | "hero_text"
  | "subtext"
  | "glass_card"
  | "image_node"
  | "kpi_block"
  | "accent_shape"
  | "divider_line"
  | "svg_chart";

export interface SlideElementBase {
  type: SlideElementType;
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface HeroTextElement extends SlideElementBase {
  type: "hero_text";
  text: string;
  fontSize?: number;   // default 40
  color?: string;      // hex with #
}

export interface SubtextElement extends SlideElementBase {
  type: "subtext";
  text: string;
  fontSize?: number;   // default 13
  color?: string;
}

export interface GlassCardElement extends SlideElementBase {
  type: "glass_card";
  title?: string;
  text: string;
}

export interface ImageNodeElement extends SlideElementBase {
  type: "image_node";
  imageUrl: string;    // Full Unsplash URL
}

export interface KpiBlockElement extends SlideElementBase {
  type: "kpi_block";
  value: string;       // e.g. "$4.2B"
  label: string;       // e.g. "إجمالي الإيرادات"
}

export interface AccentShapeElement extends SlideElementBase {
  type: "accent_shape";
  shape?: "circle" | "rect";   // default "circle"
  color?: string;              // hex with #
  opacity?: number;            // 0–1, default 0.15
}

export interface DividerLineElement extends SlideElementBase {
  type: "divider_line";
  color?: string;
}

export interface SvgChartElement extends SlideElementBase {
  type: "svg_chart";
  chartType: "bar" | "line" | "pie" | "area";
  data: { label: string; value: number; color?: string }[];
  title?: string;
}

export type SlideElement =
  | HeroTextElement
  | SubtextElement
  | GlassCardElement
  | ImageNodeElement
  | KpiBlockElement
  | AccentShapeElement
  | DividerLineElement
  | SvgChartElement;

export interface FreestyleSlide {
  slide_title: string;
  bgImagePath?: string;   // Unsplash full URL for bg
  accentColor?: string;   // Hex with #
  elements: SlideElement[];
}

export interface FreestylePresentationJSON {
  title: string;
  subtitle?: string;
  slides: FreestyleSlide[];
}

/** @deprecated Use FreestylePresentationJSON */
export type PresentationJSON = FreestylePresentationJSON;

export type DocType = "word" | "excel" | "presentation";

const SYSTEM_PROMPT_WORD = `You are an elite document architect. Convert the user's notes/ideas into a polished, well-structured document.

CRITICAL: Respond with a RAW, VALID JSON object ONLY. No markdown, no code fences, no commentary. Just JSON.

Schema:
{
  "title": "string",
  "subtitle": "string",
  "sections": [
    {
      "heading": "string",
      "subheading": "string (optional)",
      "paragraphs": ["string", ...],
      "table": { "headers": ["string"], "rows": [["string"]] }  // optional
    }
  ]
}

Rules:
- Write substantive, professional prose. Each paragraph 3-6 sentences.
- 4-8 sections typical. Include a table when it adds value (comparison, data, schedule).
- Match the document type the user asked for (academic, corporate, proposal, etc.).
- Preserve user's language (English/Arabic).`;

const SYSTEM_PROMPT_EXCEL = `You are an elite spreadsheet architect. Convert the user's requirements into a polished, well-structured spreadsheet model (Excel).

CRITICAL: Respond with a RAW, VALID JSON object ONLY. No markdown, no code fences, no commentary. Just JSON.

Schema:
{
  "title": "string (Title of the spreadsheet project)",
  "sheets": [
    {
      "name": "string (Sheet Name, e.g., 'P&L Statement', 'Inventory')",
      "headers": ["string (Header Column 1)", "string (Header Column 2)", ...],
      "rows": [
        [
          {
            "value": "string or number",
            "formula": "string (optional, e.g. '=SUM(C2:C10)' or '=B2*C2' starting with =)",
            "bold": true/false (optional),
            "bg": "string (optional, Hex color code without #, e.g. 'F3F4F6')",
            "color": "string (optional, Hex color code for text, e.g. '1E3A8A')"
          },
          ...
        ],
        ...
      ]
    }
  ]
}

Rules:
1. Formulas are extremely important. Use standard uppercase Excel formulas (like SUM, AVERAGE, IF, COUNT, etc.) to automate financial sheets, invoicing, or inventory.
2. Structure the data cleanly. Use headers properly.
3. Add summary rows at the bottom of sheets when appropriate (e.g., Total Expenses using SUM formula) and style them with bold: true and bg: "E5E7EB" or similar neutral colors.
4. Match the language requested (English/Arabic).`;

const SYSTEM_PROMPT_PRESENTATION = `You are the Lead Generative UI/UX Architect for TOLZY Flow, an elite SaaS presentation platform. Your job is to act as a human designer constructing a bespoke, ultra-premium, cinematic presentation canvas from scratch for each slide.

CRITICAL: Respond with a RAW, VALID JSON object ONLY. No markdown, no code fences, no commentary. Just JSON.
CRITICAL: Absolute prohibition of repetitive templates, pre-baked layouts, or traditional bullet-point structures. Each slide must be an individual masterpiece custom-tailored to its raw data.

### CANVAS SYSTEM SPECIFICATIONS (16:9 Widescreen)
- Total Canvas Width: 13.33 inches
- Total Canvas Height: 7.50 inches
- Target Margins: Absolute safety margin of 1.0 inch from all outer edges. Never overlap elements.

### DESIGN & VISUAL RULES
1. TYPOGRAPHY HIERARCHY (Extreme Contrast & Legibility):
   - Slide Titles (hero_text): Must be massive, commanding, and ultra-bold (44pt to 48pt). Use maximum 1 line.
   - Body/Descriptions/Subtexts: Must be clean, sharp, and highly readable (16pt to 18pt). Glass cards text must be 14pt to 16pt.
   - Always use the 'Cairo' font face rule for Arabic slides (to display text elegantly).
2. NO MORE BULLET POINTS:
   - Instead of lists, segment information into floating structural glass cards, numeric badges (e.g., '01', '02'), or visual elements like charts (svg_chart) and KPI blocks.
3. STRICT IMAGE LIMITATION:
   - Never use more than 1 or 2 "image_node" elements in a single slide. Having too many images looks unprofessional and cluttered. Only use images that directly support the slide's core message.
4. STRICT GRID COLUMN LAYOUT (No Overlaps):
   - Never overlap text elements (hero_text, subtext, glass_card) on top of image_node or svg_chart elements.
   - Arrange the slide using separate horizontal columns (e.g., Text Column on one side, Graphic/Image Column on the other).
   - If a slide has an image_node at x=7.0, w=5.3, then no text or card element should have coordinates that overlap with the x-range [7.0, 12.33]. Place the text/card elements in the x-range [1.0, 6.5] instead.
5. VISUAL adaptiveness (Context-Aware Aesthetics):
    - Analyze the slide topic. If it is financial, statistical, or contains quantitative/growth metrics, design massive KPI spotlight elements or a visual 'svg_chart'. If it is conceptual, focus on high-end negative space with short profound statements.
6. LIVE IMAGES & BACKGROUNDS:
   - For every single slide, pick an incredibly precise, context-accurate Unsplash URL for the background or internal image nodes based on the text (e.g., use dark-abstract tech images for tech slides, minimalist architecture for corporate).
   - Always use the full Unsplash URL format: https://images.unsplash.com/photo-PHOTOID?w=1400&auto=format&fit=crop&q=85

### OUTPUT SCHEMA
You must output a strict JSON object matching this precise schema:

{
  "title": "string (overall presentation title)",
  "subtitle": "string (optional subtitle)",
  "slides": [
    {
      "slide_title": "Slide Headline",
      "bgImagePath": "https://images.unsplash.com/photo-SPECIFIC_REAL_PHOTO_ID?w=1400&auto=format&fit=crop&q=85",
      "accentColor": "#HEXCODE",
      "elements": [
        {
          "type": "hero_text",
          "text": "Main striking phrase",
          "x": 1.0, "y": 1.5, "w": 7.0, "h": 2.0,
          "fontSize": 40,
          "color": "#FFFFFF"
        },
        {
          "type": "subtext",
          "text": "Supporting line that gives context without shouting.",
          "x": 1.0, "y": 3.8, "w": 7.0, "h": 0.6,
          "fontSize": 14,
          "color": "#94A3B8"
        },
        {
          "type": "glass_card",
          "title": "Card Title",
          "text": "Card body description explaining the core concept.",
          "x": 8.5, "y": 1.5, "w": 3.8, "h": 3.5
        },
        {
          "type": "image_node",
          "imageUrl": "https://images.unsplash.com/photo-SPECIFIC_REAL_PHOTO_ID?w=800&auto=format&fit=crop&q=80",
          "x": 1.0, "y": 4.5, "w": 5.0, "h": 2.0
        },
        {
          "type": "kpi_block",
          "value": "$4.2B",
          "label": "Total Revenue",
          "x": 1.0, "y": 2.0, "w": 3.0, "h": 2.0
        },
        {
          "type": "svg_chart",
          "chartType": "bar",
          "title": "Quarterly Growth",
          "data": [
            { "label": "Q1", "value": 120, "color": "#6366F1" },
            { "label": "Q2", "value": 190, "color": "#818cf8" },
            { "label": "Q3", "value": 310, "color": "#38BDF8" }
          ],
          "x": 8.5, "y": 5.2, "w": 3.8, "h": 1.8
        },
        {
          "type": "accent_shape",
          "shape": "circle",
          "color": "#6366F1",
          "opacity": 0.12,
          "x": 10.5, "y": -1.0, "w": 5.0, "h": 5.0
        },
        {
          "type": "divider_line",
          "color": "#334155",
          "x": 1.0, "y": 3.5, "w": 6.0, "h": 0.04
        }
      ]
    }
  ]
}

### ELEMENT TYPES REFERENCE
- hero_text: Main slide headline. Large (36-44pt), bold, max 1 line.
- subtext: Secondary muted description or supporting context. Small (12-15pt), muted color.
- glass_card: Floating semi-transparent card. Contains optional title + body text. Renders with glassmorphism effect.
- image_node: A cropped image region from Unsplash. Use specific photo IDs for accuracy.
- kpi_block: A large metric spotlight — huge value + small label underneath.
- svg_chart: A vector-based visual chart to display statistics, ratios, performance, and financial data. Requires 'chartType' ("bar" | "line" | "pie" | "area"), 'data' (an array of '{ label: string, value: number, color?: string }'), and optional 'title'. Highly recommended for any numeric/analytical slide!
- accent_shape: Decorative geometric shape (circle or rect). Use for depth and visual hierarchy only. Set opacity low (0.08-0.20).
- divider_line: A thin horizontal or vertical separator line.

### COORDINATE RULES
- x, y, w, h are in INCHES on a 13.33 × 7.50 canvas.
- x must be >= 1.0 and x + w must be <= 12.33 (1-inch margins).
- y must be >= 0.5 and y + h must be <= 7.2.
- Elements MUST NOT overlap each other (check all coordinates for collisions).
- Use accent_shape elements for out-of-bounds decorative bleeding (e.g., x=11.5, y=-1.0 for top-right circle bleeds).

### IMPORTANT LANGUAGE RULE
If the source text language is Arabic, write ALL element texts/titles/labels in professional, clear Arabic. JSON keys MUST always remain in English.

### SLIDE COUNT
Generate 6-10 slides total. Each slide must be visually distinct — vary the layout composition dramatically between slides (never use the same element arrangement twice).

### UNSPLASH PHOTO SELECTION
Always use real, specific Unsplash photo IDs. Here are context-matched IDs you can use:
- Dark abstract tech: photo-1518770660439-4636190af475
- Minimalist architecture: photo-1486406146926-c627a92ad1ab
- Financial/charts: photo-1611974789855-9c2a0a7236a3
- Team/people: photo-1522071820081-009f0129c71c
- Nature/sustainability: photo-1441974231531-c6227db76b6e
- AI/digital: photo-1677442135703-1787eea5ce01
- Medical/health: photo-1576091160550-2173dba999ef
- Education: photo-1523050854058-8df90110c9f1
- City/corporate: photo-1477959858617-67f85cf4f1df
- Creative/design: photo-1558618666-fcd25c85cd64
- Data/analytics: photo-1551288049-bebda4e38f71
- Real estate: photo-1560518883-ce09059eeffa
Always choose the most contextually accurate photo for the slide topic.`;

const EDIT_SYSTEM_PROMPT = `You are an elite document, spreadsheet and presentation architect. You will be given an existing document/spreadsheet/presentation in JSON format and a user's instruction.
Your task is to modify the structure or content according to the user's instruction. Keep as much of the original content, style, formulas, and structure as possible, only editing/adding/deleting elements as requested.

CRITICAL: Respond with a RAW, VALID JSON object ONLY. No markdown, no code fences, no commentary. Just JSON.

Depending on the document type, the schema MUST match exactly:
For Word documents ("word"):
{
  "title": "string",
  "subtitle": "string (optional)",
  "sections": [
    {
      "heading": "string",
      "subheading": "string (optional)",
      "paragraphs": ["string", ...],
      "table": { "headers": ["string"], "rows": [["string"]] }  // optional
    }
  ]
}

For Excel spreadsheets ("excel"):
{
  "title": "string",
  "sheets": [
    {
      "name": "string",
      "headers": ["string"],
      "rows": [
        [
          {
            "value": "string or number",
            "formula": "string (optional, e.g. '=SUM(C2:C5)' starting with =)",
            "bold": boolean (optional),
            "bg": "string (optional, Hex color, e.g., 'E5E7EB')",
            "color": "string (optional, Hex color, e.g., '1E3A8A')"
          }
        ]
      ]
    }
  ]
}

For PowerPoint presentations ("presentation"):
{
  "title": "string",
  "subtitle": "string (optional)",
  "slides": [
    {
      "slide_title": "string",
      "bgImagePath": "string (Unsplash URL)",
      "accentColor": "string (hex with #)",
      "elements": [
        {
          "type": "hero_text | subtext | glass_card | image_node | kpi_block | accent_shape | divider_line | svg_chart",
          "text": "string (for hero_text/subtext)",
          "title": "string (for glass_card/svg_chart, optional)",
          "value": "string (for kpi_block)",
          "label": "string (for kpi_block)",
          "imageUrl": "string (for image_node)",
          "chartType": "bar | line | pie | area (for svg_chart)",
          "data": "array of { label: string, value: number, color?: string } (for svg_chart)",
          "x": "number", "y": "number", "w": "number", "h": "number",
          "fontSize": "number (optional)",
          "color": "string (optional hex)",
          "shape": "circle | rect (for accent_shape)",
          "opacity": "number 0-1 (for accent_shape)"
        }
      ]
    }
  ]
}

Preserve the original language unless instructed otherwise.`;

export const generateDocument = createServerFn({ method: "POST" })
  .inputValidator((d: { prompt: string; type?: DocType }) => {
    if (!d?.prompt || typeof d.prompt !== "string" || d.prompt.length < 3) {
      throw new Error("Prompt is required");
    }
    return { prompt: d.prompt.slice(0, 8000), type: d.type || "word" };
  })
  .handler(async ({ data }) => {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) throw new Error("GEMINI_API_KEY غير مُعدّ");

    const model = "gemini-2.5-flash-lite";
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    const systemPrompt =
      data.type === "excel"
        ? SYSTEM_PROMPT_EXCEL
        : data.type === "presentation"
        ? SYSTEM_PROMPT_PRESENTATION
        : SYSTEM_PROMPT_WORD;

    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        systemInstruction: { parts: [{ text: systemPrompt }] },
        contents: [{ role: "user", parts: [{ text: data.prompt }] }],
        generationConfig: { responseMimeType: "application/json" },
      }),
    });

    if (!res.ok) {
      const text = await res.text();
      console.error("Gemini API error", res.status, text);
      if (res.status === 429) throw new Error("تم تجاوز حد الاستخدام. حاول لاحقًا.");
      throw new Error(`Gemini ${res.status}: ${text.slice(0, 300)}`);
    }

    const json = await res.json();
    const content: string =
      json?.candidates?.[0]?.content?.parts?.map((p: { text?: string }) => p.text ?? "").join("") ?? "";
    const cleaned = content.trim().replace(/^```json\s*|\s*```$/g, "");

    let parsed: any;
    try {
      parsed = JSON.parse(cleaned);
    } catch {
      throw new Error("استجابة Gemini ليست JSON صالحًا");
    }
    // For presentation, validate slides array exists
    if (data.type === "presentation") {
      if (!parsed?.slides || !Array.isArray(parsed.slides) || parsed.slides.length === 0) {
        throw new Error("استجابة Gemini ينقصها مصفوفة الشرائح");
      }
      // Inject a fallback title if missing
      if (!parsed.title) parsed.title = parsed.slides[0]?.slide_title ?? "TOLZY Presentation";
    } else {
      if (!parsed?.title) {
        throw new Error("استجابة Gemini ينقصها حقول مطلوبة");
      }
    }
    return parsed;
  });

export const editDocument = createServerFn({ method: "POST" })
  .inputValidator((d: { currentDocument: any; instruction: string; type?: DocType }) => {
    if (!d?.currentDocument) throw new Error("Current document is required");
    if (!d?.instruction || typeof d.instruction !== "string") throw new Error("Instruction is required");
    return { currentDocument: d.currentDocument, instruction: d.instruction, type: d.type || "word" };
  })
  .handler(async ({ data }) => {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) throw new Error("GEMINI_API_KEY غير مُعدّ");

    const model = "gemini-2.5-flash-lite";
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        systemInstruction: { parts: [{ text: EDIT_SYSTEM_PROMPT }] },
        contents: [
          {
            role: "user",
            parts: [
              { text: `Document Type: ${data.type}\nCurrent Document JSON:\n${JSON.stringify(data.currentDocument, null, 2)}` },
              { text: `User instruction: ${data.instruction}` }
            ]
          }
        ],
        generationConfig: { responseMimeType: "application/json" },
      }),
    });

    if (!res.ok) {
      const text = await res.text();
      console.error("Gemini API error during edit", res.status, text);
      if (res.status === 429) throw new Error("تم تجاوز حد الاستخدام. حاول لاحقًا.");
      throw new Error(`Gemini ${res.status}: ${text.slice(0, 300)}`);
    }

    const json = await res.json();
    const content: string =
      json?.candidates?.[0]?.content?.parts?.map((p: { text?: string }) => p.text ?? "").join("") ?? "";
    const cleaned = content.trim().replace(/^```json\s*|\s*```$/g, "");

    let parsed: any;
    try {
      parsed = JSON.parse(cleaned);
    } catch {
      throw new Error("استجابة Gemini للتعديل ليست JSON صالحًا");
    }
    if (data.type === "presentation") {
      if (!parsed?.slides || !Array.isArray(parsed.slides)) {
        throw new Error("استجابة Gemini للتعديل ينقصها مصفوفة الشرائح");
      }
      if (!parsed.title) parsed.title = parsed.slides[0]?.slide_title ?? "TOLZY Presentation";
    } else {
      if (!parsed?.title) {
        throw new Error("استجابة Gemini للتعديل ينقصها حقول مطلوبة");
      }
    }
    return parsed;
  });
