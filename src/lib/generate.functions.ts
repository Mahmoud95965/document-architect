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

const SYSTEM_PROMPT_WORD = `You are an elite, world-class executive document architect for TOLZY Flow.
Your job is to transform user prompts into extremely comprehensive, highly realistic, professional, and detailed documents (reports, proposals, specifications, research papers, legal frameworks, business plans).

CRITICAL INSTRUCTIONS:
1. OUTPUT ONLY A RAW, VALID JSON OBJECT. Do NOT include markdown code fences (no \`\`\`json), no commentary, no preambles.
2. DENSITY & RICHNESS:
   - Provide extensive, highly realistic, substantive content.
   - Generate 6 to 10 detailed sections.
   - Every paragraph must be 4 to 8 long, well-written sentences filled with authentic domain terminology, precise metrics, industry best practices, and thorough analysis. Avoid short or lazy generic placeholders.
   - Include high-value, structured data tables inside sections whenever relevant (e.g., financial breakdowns, project timelines, comparative metrics, feature matrices).
3. LANGUAGE: Match the user's language (Arabic or English) with flawless grammar and executive prose.

JSON Schema:
{
  "title": "string (Main document title)",
  "subtitle": "string (Executive subtitle or description)",
  "sections": [
    {
      "heading": "string (Section Heading)",
      "subheading": "string (Optional Subheading)",
      "paragraphs": [
        "string (Long substantive paragraph 1 with deep details...)",
        "string (Long substantive paragraph 2 with actionable data...)"
      ],
      "table": {
        "headers": ["string", "string", "string"],
        "rows": [
          ["string", "string", "string"],
          ["string", "string", "string"]
        ]
      }
    }
  ]
}`;

const SYSTEM_PROMPT_EXCEL = `You are a Chief Financial Officer (CFO) and Principal Data Architect for TOLZY Flow.
Your job is to convert user requests into ultra-realistic, multi-sheet financial models, budget statements, operational workbooks, and analytical spreadsheets (Excel).

CRITICAL INSTRUCTIONS:
1. OUTPUT ONLY A RAW, VALID JSON OBJECT. Do NOT include markdown code fences (no \`\`\`json), no commentary.
2. RICHNESS & REALISM:
   - Create 2 to 4 distinct, fully-populated worksheet sheets (e.g., 'ملخص تنفيذي Executive Summary', 'التحليل الشهري Monthly Data', 'المؤشرات والنسب KPIs', 'تفاصيل الحسابات Detailed Accounts').
   - Provide realistic numerical data, real-world expense categories, revenue streams, and unit economics.
3. EXCEL FORMULAS ARE MANDATORY:
   - Use standard uppercase Excel formulas starting with '=' (e.g., '=SUM(C3:C12)', '=AVERAGE(D3:D12)', '=IF(C15>0, "Profitable", "Deficit")', '=B5*C5', '=SUM(E3:E20)').
4. STYLING & FORMATTING (Hex colors without '#'):
   - Header Row: Use bold text (bold: true), dark background (bg: "1E3A8A" or "065F46"), white text (color: "FFFFFF").
   - Total/Summary Row: Bold text (bold: true), highlight background (bg: "F3F4F6" or "E0E7FF"), dark text (color: "1E1B4B").
5. LANGUAGE: Match user's requested language (Arabic or English).

JSON Schema:
{
  "title": "string (Project Title)",
  "sheets": [
    {
      "name": "string (Sheet Name)",
      "headers": ["string", "string", "string"],
      "rows": [
        [
          {
            "value": "string or number",
            "formula": "string (optional, uppercase formula e.g. '=SUM(B2:B10)')",
            "bold": true/false,
            "bg": "string (optional hex without # e.g. '1E3A8A')",
            "color": "string (optional hex without # e.g. 'FFFFFF')"
          }
        ]
      ]
    }
  ]
}`;

const SYSTEM_PROMPT_PRESENTATION = `You are the Lead Visual Experience Architect for TOLZY Flow. Your job is to construct a cinematic, bespoke presentation deck from scratch for each slide.

CRITICAL INSTRUCTIONS:
1. OUTPUT ONLY A RAW, VALID JSON OBJECT. Do NOT include markdown code fences (no \`\`\`json), no commentary.
2. SLIDE DECK DENSITY: Generate 7 to 12 visually distinct widescreen (16:9) slides.
3. COMPOSITION & NO BULLET POINTS:
   - Never use bulleted lists.
   - Use dynamic visual elements: 'hero_text', 'subtext', 'glass_card', 'kpi_block', 'svg_chart', 'image_node', 'accent_shape', 'divider_line'.
   - Include 'svg_chart' for data slides with chartType ("bar" | "line" | "pie" | "area") and data array.
   - Include 'kpi_block' for quantitative spotlights ($5.4M, 98.2%, +140%).
4. COORDINATES & NON-OVERLAPPING GRID (13.33 × 7.50 Canvas):
   - Safety margins: x >= 1.0, x + w <= 12.33; y >= 0.5, y + h <= 7.2.
   - Never overlap text or cards over image_node or svg_chart elements.
5. UNSPLASH PHOTOS: Use real Unsplash photo URLs:
   - Tech/AI: https://images.unsplash.com/photo-1518770660439-4636190af475?w=1400&auto=format&fit=crop&q=85
   - Business/Corporate: https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?w=1400&auto=format&fit=crop&q=85
   - Finance/Data: https://images.unsplash.com/photo-1611974789855-9c2a0a7236a3?w=1400&auto=format&fit=crop&q=85
   - Team/Office: https://images.unsplash.com/photo-1522071820081-009f0129c71c?w=1400&auto=format&fit=crop&q=85
6. LANGUAGE: Match Arabic or English as requested by user.

JSON Schema:
{
  "title": "string",
  "subtitle": "string",
  "slides": [
    {
      "slide_title": "string",
      "bgImagePath": "string (Unsplash URL)",
      "accentColor": "#HEXCODE",
      "elements": [
        {
          "type": "hero_text | subtext | glass_card | image_node | kpi_block | accent_shape | divider_line | svg_chart",
          "text": "string",
          "title": "string",
          "value": "string",
          "label": "string",
          "imageUrl": "string",
          "chartType": "bar | line | pie | area",
          "data": [{ "label": "string", "value": 100, "color": "#HEX" }],
          "x": 1.0, "y": 1.5, "w": 6.0, "h": 2.0,
          "fontSize": 40, "color": "#FFFFFF",
          "shape": "circle | rect", "opacity": 0.15
        }
      ]
    }
  ]
}`;

const EDIT_SYSTEM_PROMPT = `You are an elite document, spreadsheet, and presentation architect for TOLZY Flow.
You will receive an existing document/spreadsheet/presentation JSON and a user instruction.
Modify and refine the JSON according to the instruction while maintaining maximum content density, professional structure, valid formulas, and high visual quality.

CRITICAL: Respond with a RAW, VALID JSON object ONLY. No markdown fences, no commentary.`;

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

// ─── DAHL INFERENCE API + GEMINI FALLBACK WRAPPER ───

const DAHL_API_KEY = "dahl_GDrj43To46oAjpGbvFZ3ATEtc439GE63h";
const DAHL_ENDPOINT = "https://inference.dahl.global/v1/chat/completions";
const DAHL_MODEL = "moonshotai/Kimi-K2.6";

async function fetchCompletionFromAI(systemPrompt: string, userPrompt: string): Promise<any> {
  // 1. Primary Provider: Dahl Inference API (Kimi-K2.6)
  try {
    const dahlRes = await fetch(DAHL_ENDPOINT, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${DAHL_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: DAHL_MODEL,
        response_format: { type: "json_object" },
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userPrompt },
        ],
        temperature: 0.3,
        max_tokens: 8192,
      }),
    });

    if (dahlRes.ok) {
      const json = await dahlRes.json();
      const content = json?.choices?.[0]?.message?.content;
      if (content && typeof content === "string" && content.trim().length > 0) {
        try {
          return cleanAndParseJSON(content);
        } catch (parseErr) {
          console.warn("Failed to parse Dahl output, falling back to Gemini:", parseErr);
        }
      }
    } else {
      console.warn("Dahl API response status:", dahlRes.status);
    }
  } catch (err) {
    console.warn("Dahl API connection error, falling back to Gemini:", err);
  }

  // 2. Secondary Backup Provider: Gemini API (gemini-2.5-flash-lite)
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error("تعذر الاتصال بالمزود الرئيسي وليس هناك مفتاح احتياطي لـ Gemini");
  }

  const model = "gemini-2.5-flash-lite";
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      systemInstruction: { parts: [{ text: systemPrompt }] },
      contents: [{ role: "user", parts: [{ text: userPrompt }] }],
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
  const content =
    json?.candidates?.[0]?.content?.parts?.map((p: { text?: string }) => p.text ?? "").join("") ?? "";
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
