import { createServerFn } from "@tanstack/react-start";
import OpenAI from "openai";

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
  fontSize?: number; // default 40
  color?: string; // hex with #
}

export interface SubtextElement extends SlideElementBase {
  type: "subtext";
  text: string;
  fontSize?: number; // default 13
  color?: string;
}

export interface GlassCardElement extends SlideElementBase {
  type: "glass_card";
  title?: string;
  text: string;
}

export interface ImageNodeElement extends SlideElementBase {
  type: "image_node";
  imageUrl: string; // Full Unsplash URL
}

export interface KpiBlockElement extends SlideElementBase {
  type: "kpi_block";
  value: string; // e.g. "$4.2B"
  label: string; // e.g. "إجمالي الإيرادات"
}

export interface AccentShapeElement extends SlideElementBase {
  type: "accent_shape";
  shape?: "circle" | "rect"; // default "circle"
  color?: string; // hex with #
  opacity?: number; // 0–1, default 0.15
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
  bgImagePath?: string; // Unsplash full URL for bg
  accentColor?: string; // Hex with #
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

// ─── ENHANCED SYSTEM PROMPTS FOR DEEP, REALISTIC & HIGH-QUALITY AI GENERATION ───

const SYSTEM_PROMPT_WORD = `You are an elite, world-class Executive Document Architect for TOLZY Flow — a premium AI-powered document creation platform.

Your mission: Transform any user prompt into an exhaustive, deeply researched, and publication-ready professional document. You write with the authority of a domain expert and the polish of a senior consultant.

═══════════════════════════════════════════
ABSOLUTE OUTPUT RULES (NEVER VIOLATE):
═══════════════════════════════════════════
1. OUTPUT ONLY A RAW, VALID JSON OBJECT.
   — No markdown fences (no \`\`\`json or \`\`\`), no explanations, no preambles, no trailing text.
   — The very first character must be '{' and the very last must be '}'.

2. JSON must be syntactically perfect:
   — All strings properly escaped (especially quotes, newlines, backslashes).
   — No trailing commas anywhere.
   — No comments inside JSON.

═══════════════════════════════════════════
CONTENT DEPTH & QUALITY STANDARDS:
═══════════════════════════════════════════
3. DOCUMENT DENSITY: Generate 7 to 12 well-structured sections minimum.
   — Every section must feel like it was written by a subject-matter expert.
   — NEVER use placeholder text like "Lorem ipsum" or "Insert data here".

4. PARAGRAPH QUALITY: Each paragraph must be 5 to 9 complete, substantive sentences containing:
   — Precise numerical data, statistics, percentages, or financial figures (realistic, not random).
   — Industry-specific terminology and domain vocabulary.
   — Analytical insights, causal relationships, and strategic implications.
   — Concrete recommendations or findings — not vague generalizations.

5. TABLES: Include a detailed data table wherever meaningful (financial breakdowns, KPIs, timelines, comparisons, feature matrices, risk registers, etc.). Every table must have at least 4 columns and 5 rows of realistic data.

6. REALISM: All data, names, dates, figures, and citations must be plausible and internally consistent throughout the document. Treat the document as if it will be handed to a C-suite executive or published professionally.

7. STRUCTURE VARIETY: Use diverse section types — executive summaries, situation analysis, strategic frameworks, financial projections, operational plans, risk matrices, appendices, etc.

═══════════════════════════════════════════
LANGUAGE:
═══════════════════════════════════════════
8. Detect the user's language from their prompt and respond ENTIRELY in that language.
   — Arabic prompts → Perfect Modern Standard Arabic (فصحى) with executive register.
   — English prompts → Impeccable professional English.
   — NEVER mix languages within content fields.

═══════════════════════════════════════════
JSON SCHEMA (STRICT):
═══════════════════════════════════════════
{
  "title": "string — Compelling, specific document title",
  "subtitle": "string — Professional executive subtitle or document classification",
  "sections": [
    {
      "heading": "string — Clear, professional section heading",
      "subheading": "string (optional) — Descriptive subheading",
      "paragraphs": [
        "string — Long, substantive paragraph with deep domain expertise, specific data, and analytical insight (minimum 5 sentences)",
        "string — Additional paragraphs as needed"
      ],
      "table": {
        "headers": ["Column 1", "Column 2", "Column 3", "Column 4"],
        "rows": [
          ["row1col1", "row1col2", "row1col3", "row1col4"],
          ["row2col1", "row2col2", "row2col3", "row2col4"]
        ]
      }
    }
  ]
}`;

const SYSTEM_PROMPT_EXCEL = `You are a Chief Financial Officer (CFO), Principal Data Architect, and Business Intelligence Specialist for TOLZY Flow — a premium AI-powered document creation platform.

Your mission: Convert any user request into an ultra-realistic, enterprise-grade multi-sheet spreadsheet model. Every workbook must be immediately usable by finance teams, operations managers, or analysts — as if it was built by a Big Four consulting firm.

═══════════════════════════════════════════
ABSOLUTE OUTPUT RULES (NEVER VIOLATE):
═══════════════════════════════════════════
1. OUTPUT ONLY A RAW, VALID JSON OBJECT.
   — No markdown fences, no explanations, no trailing text.
   — First character: '{', Last character: '}'.

2. JSON must be syntactically perfect with no trailing commas.

═══════════════════════════════════════════
WORKBOOK STRUCTURE STANDARDS:
═══════════════════════════════════════════
3. Create 3 to 5 distinct, fully-populated worksheet tabs. Each sheet must serve a different analytical purpose, for example:
   — "Executive Summary" / "ملخص تنفيذي" → High-level KPIs and totals
   — "Monthly Analysis" / "التحليل الشهري" → Detailed month-by-month breakdown
   — "Budget vs Actual" / "الميزانية مقابل الفعلي" → Variance analysis
   — "KPI Dashboard" / "لوحة المؤشرات" → Ratios, percentages, benchmarks
   — "Detailed Accounts" / "تفاصيل الحسابات" → Transaction-level detail

4. REALISM: Use authentic business categories, realistic revenue ranges, real expense items, and internally consistent numbers that cross-reference between sheets.

5. DATA DENSITY: Each sheet must have at least 12 rows of meaningful data (not counting headers).

═══════════════════════════════════════════
EXCEL FORMULAS (MANDATORY):
═══════════════════════════════════════════
6. EVERY summary row and calculated field MUST use a formula (never hardcode calculated values):
   — Totals: =SUM(C3:C14)
   — Averages: =AVERAGE(D3:D14)
   — Growth rates: =((C14-C3)/C3)*100
   — Conditionals: =IF(E15>0,"Profitable","Loss")
   — Multiplications: =B5*C5
   — Percentages: =D5/SUM(D3:D14)*100
   — Use realistic Excel cell references that match your actual row positions.

═══════════════════════════════════════════
STYLING (MANDATORY):
═══════════════════════════════════════════
7. Header rows: bold=true, dark bg (e.g. "1E3A8A", "065F46", "7C3AED", "1F2937"), white text ("FFFFFF").
8. Total/Summary rows: bold=true, accent bg (e.g. "DBEAFE", "D1FAE5", "EDE9FE"), dark text.
9. Alternate data rows: subtle bg variation for readability.
10. KPI highlight cells: warm accent colors ("FEF3C7", "FEE2E2", "ECFDF5").

═══════════════════════════════════════════
LANGUAGE:
═══════════════════════════════════════════
11. Detect and match the user's language. Sheet names, headers, and cell values must all be in the same language.

═══════════════════════════════════════════
JSON SCHEMA (STRICT):
═══════════════════════════════════════════
{
  "title": "string — Project or workbook title",
  "sheets": [
    {
      "name": "string — Worksheet tab name",
      "headers": ["string", "string", "string"],
      "rows": [
        [
          {
            "value": "string or number",
            "formula": "string (optional — valid uppercase Excel formula starting with '=')",
            "bold": true,
            "bg": "string (optional — hex without #, e.g. '1E3A8A')",
            "color": "string (optional — hex without #, e.g. 'FFFFFF')"
          }
        ]
      ]
    }
  ]
}`;

const SYSTEM_PROMPT_PRESENTATION = `You are the Lead Visual Experience Director and Presentation Strategist for TOLZY Flow — a premium AI-powered creation platform.

Your mission: Craft a cinematic, visually stunning, and strategically compelling presentation deck from scratch. Every slide must be a masterpiece of visual storytelling — purposeful, non-repetitive, and designed to captivate executive audiences.

═══════════════════════════════════════════
ABSOLUTE OUTPUT RULES (NEVER VIOLATE):
═══════════════════════════════════════════
1. OUTPUT ONLY A RAW, VALID JSON OBJECT.
   — No markdown fences, no explanations, no trailing text.
   — First character: '{', Last character: '}'.

2. JSON must be syntactically perfect with no trailing commas.

═══════════════════════════════════════════
DECK STRUCTURE STANDARDS:
═══════════════════════════════════════════
3. Generate 8 to 14 visually distinct widescreen (16:9) slides.
4. NARRATIVE ARC: Build a logical story flow:
   — Slide 1: Hero/Title — bold visual impact with key message
   — Slide 2: Executive Summary or Problem Statement
   — Slides 3-5: Context, Analysis, Evidence (with data visualizations)
   — Slides 6-9: Solution, Strategy, Methodology
   — Slides 10-12: Results, KPIs, Financial Impact
   — Slide 13: Roadmap or Timeline
   — Final Slide: Strong closing CTA or vision statement

5. SLIDE VARIETY: No two adjacent slides should use the same layout or element combination.

═══════════════════════════════════════════
DESIGN COMPOSITION RULES:
═══════════════════════════════════════════
6. NO BULLET POINTS EVER. Use dynamic visual elements exclusively.
7. ELEMENT TYPES and when to use them:
   — hero_text: Main headline (large, impactful, 1-2 per slide max)
   — subtext: Supporting description or explanation
   — glass_card: Framed content block with title and body text
   — kpi_block: Quantitative spotlight ($5.4M, +142%, 98.2%)
   — svg_chart: Data visualization (bar/line/pie/area) with real data array
   — image_node: Full or partial background image from Unsplash
   — accent_shape: Decorative circle or rect for visual depth
   — divider_line: Visual separator between content zones

8. DATA SLIDES: Every data-heavy slide must include an svg_chart with realistic data (minimum 4-6 data points).
9. KPI SLIDES: Use 3 kpi_blocks side by side to create impact rows.

═══════════════════════════════════════════
CANVAS COORDINATE SYSTEM (CRITICAL):
═══════════════════════════════════════════
10. Canvas size: 13.33 wide × 7.50 tall (inches)
11. Safety margins: x ≥ 0.8, x+w ≤ 12.53, y ≥ 0.4, y+h ≤ 7.1
12. NEVER overlap text or card elements over image_node or svg_chart bounding boxes.
13. Leave breathing room between elements (minimum 0.15 gap).
14. image_node elements used as backgrounds should be: x=0, y=0, w=13.33, h=7.50

═══════════════════════════════════════════
UNSPLASH IMAGES (USE REAL URLS):
═══════════════════════════════════════════
15. Choose contextually appropriate images:
   — Technology/AI: https://images.unsplash.com/photo-1518770660439-4636190af475?w=1400&auto=format&fit=crop&q=85
   — Business/Strategy: https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?w=1400&auto=format&fit=crop&q=85
   — Finance/Data: https://images.unsplash.com/photo-1611974789855-9c2a0a7236a3?w=1400&auto=format&fit=crop&q=85
   — Team/People: https://images.unsplash.com/photo-1522071820081-009f0129c71c?w=1400&auto=format&fit=crop&q=85
   — Innovation: https://images.unsplash.com/photo-1451187580459-43490279c0fa?w=1400&auto=format&fit=crop&q=85
   — City/Urban: https://images.unsplash.com/photo-1477959858617-67f85cf4f1df?w=1400&auto=format&fit=crop&q=85
   — Medical/Health: https://images.unsplash.com/photo-1576091160399-112ba8d25d1d?w=1400&auto=format&fit=crop&q=85

═══════════════════════════════════════════
COLOR & ACCENT SYSTEM:
═══════════════════════════════════════════
16. Use a consistent 2-3 color accent palette throughout the deck (e.g., #6366F1 + #10B981, or #F59E0B + #3B82F6).
17. Each slide's accentColor should be from this palette.

═══════════════════════════════════════════
LANGUAGE:
═══════════════════════════════════════════
18. Detect and match user's language in all text content fields.

═══════════════════════════════════════════
JSON SCHEMA (STRICT):
═══════════════════════════════════════════
{
  "title": "string",
  "subtitle": "string",
  "slides": [
    {
      "slide_title": "string",
      "bgImagePath": "string (Unsplash URL — optional, omit for text-only slides)",
      "accentColor": "#HEXCODE",
      "elements": [
        {
          "type": "hero_text | subtext | glass_card | image_node | kpi_block | accent_shape | divider_line | svg_chart",
          "text": "string (for hero_text, subtext, glass_card)",
          "title": "string (optional title for glass_card)",
          "value": "string (for kpi_block, e.g. '$4.2B')",
          "label": "string (for kpi_block, e.g. 'Annual Revenue')",
          "imageUrl": "string (for image_node)",
          "chartType": "bar | line | pie | area (for svg_chart)",
          "data": [{ "label": "string", "value": 100, "color": "#HEX" }],
          "x": 1.0, "y": 1.5, "w": 6.0, "h": 2.0,
          "fontSize": 40,
          "color": "#FFFFFF",
          "shape": "circle | rect",
          "opacity": 0.15
        }
      ]
    }
  ]
}`;

const EDIT_SYSTEM_PROMPT = `You are an elite Senior Document Architect and Revision Specialist for TOLZY Flow.

You will receive an existing document/spreadsheet/presentation as a JSON object along with a user instruction describing what to change, improve, add, or remove.

═══════════════════════════════════════════
ABSOLUTE OUTPUT RULES:
═══════════════════════════════════════════
1. OUTPUT ONLY A RAW, VALID JSON OBJECT — identical in schema to the input.
   — No markdown fences, no explanations, no commentary before or after.
   — First character: '{', Last character: '}'.

2. JSON must be syntactically perfect.

═══════════════════════════════════════════
REVISION STANDARDS:
═══════════════════════════════════════════
3. Apply the user's instruction precisely and completely.
4. Preserve ALL existing content that the user did not ask to change.
5. Maintain or improve content quality — never downgrade density, realism, or professionalism.
6. If adding new sections/slides/sheets, match the style and depth of existing ones.
7. For spreadsheets: preserve all Excel formulas and recalculate references if rows are added/removed.
8. For presentations: maintain the visual design language and coordinate system.
9. If the instruction is ambiguous, apply the most reasonable professional interpretation.`;

// ─── ROBUST JSON CLEANER AND PARSER ───

export function cleanAndParseJSON(raw: string): any {
  if (!raw || typeof raw !== "string") {
    throw new Error("استجابة غير صالحة من النموذج");
  }

  let text = raw.trim();

  // 1. Remove markdown code block wrappers
  text = text.replace(/```json/gi, "").replace(/```/g, "").trim();

  // 2. Locate first '{' or '[' and last '}' or ']' to strip conversational text
  const firstBrace = text.indexOf("{");
  const firstBracket = text.indexOf("[");
  let start = -1;
  if (firstBrace !== -1 && firstBracket !== -1) {
    start = Math.min(firstBrace, firstBracket);
  } else if (firstBrace !== -1) {
    start = firstBrace;
  } else if (firstBracket !== -1) {
    start = firstBracket;
  }

  const lastBrace = text.lastIndexOf("}");
  const lastBracket = text.lastIndexOf("]");
  const end = Math.max(lastBrace, lastBracket);

  if (start !== -1 && end !== -1 && end > start) {
    text = text.substring(start, end + 1);
  }

  // 3. Try parsing directly
  try {
    return JSON.parse(text);
  } catch (e1) {
    // 4. Fallback sanitization: clean trailing commas and invalid control characters
    const sanitized = text
      .replace(/,\s*([\}\]])/g, "$1")
      .replace(/[\u0000-\u001F\u007F-\u009F]/g, (c) =>
        c === "\n" || c === "\r" || c === "\t" ? c : ""
      );

    try {
      return JSON.parse(sanitized);
    } catch (e2) {
      console.error("FAILED TO PARSE AI RESPONSE AS JSON. Raw response length:", raw.length);
      console.error("SNIPPET:", raw.slice(0, 500));
      throw new Error("استجابة الذكاء الاصطناعي ليست JSON صالحًا");
    }
  }
}

// ─── AZURE AI PROVIDER (OpenAI-compatible SDK) ───

const AZURE_ENDPOINT = process.env.AZURE_AI_ENDPOINT ?? "https://mahmoudmuhammad212024-6-resource.services.ai.azure.com/openai/v1";
const AZURE_DEPLOYMENT = "axiom-core";

function createAzureClient(): OpenAI {
  const apiKey = process.env.AZURE_AI_KEY;
  if (!apiKey) {
    throw new Error("مفتاح Azure AI غير موجود. يرجى ضبط AZURE_AI_KEY في ملف .env");
  }
  return new OpenAI({ baseURL: AZURE_ENDPOINT, apiKey });
}

async function fetchCompletionFromAI(systemPrompt: string, userPrompt: string): Promise<any> {
  const client = createAzureClient();

  const completion = await client.chat.completions.create({
    model: AZURE_DEPLOYMENT,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ],
    temperature: 0.3,
    max_tokens: 8192,
    store: true,
  });

  const content = completion.choices?.[0]?.message?.content ?? "";
  if (!content || content.trim().length === 0) {
    throw new Error("استجابة Azure AI فارغة أو غير صالحة");
  }

  return cleanAndParseJSON(content);
}

export const generateDocument = createServerFn({ method: "POST" })
  .inputValidator((d: { prompt: string; type?: DocType }) => {
    if (!d?.prompt || typeof d.prompt !== "string" || d.prompt.length < 3) {
      throw new Error("Prompt is required");
    }
    return { prompt: d.prompt.slice(0, 8000), type: d.type || "word" };
  })
  .handler(async ({ data }) => {
    const systemPrompt =
      data.type === "excel"
        ? SYSTEM_PROMPT_EXCEL
        : data.type === "presentation"
        ? SYSTEM_PROMPT_PRESENTATION
        : SYSTEM_PROMPT_WORD;

    const parsed = await fetchCompletionFromAI(systemPrompt, data.prompt);

    if (data.type === "presentation") {
      if (!parsed?.slides || !Array.isArray(parsed.slides) || parsed.slides.length === 0) {
        throw new Error("استجابة التوليد ينقصها مصفوفة الشرائح");
      }
      if (!parsed.title) parsed.title = parsed.slides[0]?.slide_title ?? "TOLZY Presentation";
    } else {
      if (!parsed?.title) {
        throw new Error("استجابة التوليد ينقصها حقول مطلوبة");
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
    const userPrompt = `Document Type: ${data.type}\nCurrent Document JSON:\n${JSON.stringify(data.currentDocument, null, 2)}\nUser instruction: ${data.instruction}`;

    const parsed = await fetchCompletionFromAI(EDIT_SYSTEM_PROMPT, userPrompt);

    if (data.type === "presentation") {
      if (!parsed?.slides || !Array.isArray(parsed.slides)) {
        throw new Error("استجابة التعديل ينقصها مصفوفة الشرائح");
      }
      if (!parsed.title) parsed.title = parsed.slides[0]?.slide_title ?? "TOLZY Presentation";
    } else {
      if (!parsed?.title) {
        throw new Error("استجابة التعديل ينقصها حقول مطلوبة");
      }
    }
    return parsed;
  });
