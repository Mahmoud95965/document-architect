import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useServerFn } from "@tanstack/react-start";
import { useState, useEffect } from "react";
import { motion, AnimatePresence } from "framer-motion";
import {
  FileText, Sparkles, Download, ArrowRight, Trash2, Plus,
  ChevronLeft, Edit, Eye, PlusCircle, MinusCircle, Save,
  Loader2, FileSpreadsheet, Table2, LayoutTemplate,
  Presentation, Zap, History, Clock, FolderOpen,
  Cpu, ChevronRight, LogOut, Lock, ExternalLink, User, Moon, Sun
} from "lucide-react";
import { useAuth } from "@/hooks/useAuth";
import { Toaster, toast } from "sonner";
import {
  generateDocument,
  editDocument,
  type ExcelSheet,
  type ExcelCell,
  type FreestyleSlide,
} from "@/lib/generate.functions";
import { generateAndDownloadDocx } from "@/lib/buildDocx";
import { generateAndDownloadExcel } from "@/lib/buildExcel";
import { generateAndDownloadPptx, getUnsplashImageUrl } from "@/lib/buildPptx";
import { getProjects, saveProject, deleteProject, type SavedProject } from "@/lib/projects";
import { isProjectArabic } from "@/lib/lang";

export const Route = createFileRoute("/")({
  component: Index,
  head: () => ({
    meta: [
      { title: "TOLZY Flow — AI Document & Presentation Architect" },
      {
        name: "description",
        content:
          "Turn raw notes into Word documents, Excel spreadsheets and PowerPoint presentations. AI-crafted structure, instant download.",
      },
    ],
  }),
});

const WORD_EXAMPLES = [
  "An academic research paper outline on poultry production efficiency",
  "A corporate project proposal for a fintech onboarding revamp",
  "A product spec for a real-time collaboration whiteboard",
];

const EXCEL_EXAMPLES = [
  "لوحة أرباح وخسائر لشركة استيراد وتصدير لمدة 6 أشهر مع معادلات تلقائية",
  "فاتورة احترافية لشحنة بضاعة تشمل حساب الضريبة والخصم والإجمالي",
  "جدول مخزون ذكي مع تتبع الكميات وتواريخ الوصول وتنبيهات المخزون المنخفض",
];

const PRESENTATION_EXAMPLES = [
  "عرض تقديمي لمشروع عقاري سكني فاخر يشمل الأهداف والموقع وتفاصيل الاستثمار",
  "Pitch deck for a healthcare telemedicine app startup including problem, solution, market size",
  "ملخص استراتيجية التسويق الرقمي لعام 2026 شاملة القنوات والميزانية والنتائج المتوقعة",
];

type DocMode = "word" | "excel" | "presentation";

function colLabel(i: number): string {
  let s = "";
  for (; i >= 0; i = Math.floor(i / 26) - 1)
    s = String.fromCharCode(65 + (i % 26)) + s;
  return s;
}

// ── Tiny helpers ──────────────────────────────────────────
function ModeIcon({ mode, size = "h-4 w-4" }: { mode: DocMode; size?: string }) {
  if (mode === "excel") return <FileSpreadsheet className={`${size} text-emerald-400`} />;
  if (mode === "presentation") return <Presentation className={`${size} text-orange-400`} />;
  return <FileText className={`${size} text-indigo-400`} />;
}

function TypeBadge({ type }: { type: string }) {
  const cfgMap: Record<string, { cls: string; label: string }> = {
    word:         { cls: "badge-word",  label: "Word" },
    excel:        { cls: "badge-excel", label: "Excel" },
    presentation: { cls: "badge-ppt",   label: "PowerPoint" },
  };
  const cfg = cfgMap[type] ?? { cls: "badge-word", label: type || "Doc" };
  return (
    <span className={`text-[10px] font-bold px-2.5 py-1 rounded-full ${cfg.cls}`}>
      {cfg.label}
    </span>
  );
}

// ── Debounced input wrappers to fix Typing Input Lag ───────
function DebouncedInput({
  value,
  onChange,
  className,
  style,
  placeholder,
  ...props
}: {
  value: string;
  onChange: (val: string) => void;
  className?: string;
  style?: React.CSSProperties;
  placeholder?: string;
  [key: string]: any;
}) {
  const [localVal, setLocalVal] = useState(value ?? "");

  useEffect(() => {
    setLocalVal(value ?? "");
  }, [value]);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setLocalVal(e.target.value);
  };

  const handleBlur = () => {
    if (localVal !== value) {
      onChange(localVal);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === "Enter") {
      if (localVal !== value) {
        onChange(localVal);
      }
      e.currentTarget.blur();
    }
  };

  return (
    <input
      type="text"
      value={localVal}
      onChange={handleChange}
      onBlur={handleBlur}
      onKeyDown={handleKeyDown}
      className={className}
      style={style}
      placeholder={placeholder}
      {...props}
    />
  );
}

function DebouncedTextarea({
  value,
  onChange,
  className,
  style,
  placeholder,
  rows,
  ...props
}: {
  value: string;
  onChange: (val: string) => void;
  className?: string;
  style?: React.CSSProperties;
  placeholder?: string;
  rows?: number;
  [key: string]: any;
}) {
  const [localVal, setLocalVal] = useState(value ?? "");

  useEffect(() => {
    setLocalVal(value ?? "");
  }, [value]);

  const handleChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setLocalVal(e.target.value);
  };

  const handleBlur = () => {
    if (localVal !== value) {
      onChange(localVal);
    }
  };

  return (
    <textarea
      value={localVal}
      onChange={handleChange}
      onBlur={handleBlur}
      className={className}
      style={style}
      placeholder={placeholder}
      rows={rows}
      {...props}
    />
  );
}

// ════════════════════════════════════════════════════════════
//  INDEX COMPONENT
// ════════════════════════════════════════════════════════════
function Index() {
  const [isDark, setIsDark] = useState(true);
  
  useEffect(() => {
    if (isDark) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  }, [isDark]);
  const generate   = useServerFn(generateDocument);
  const editWithAI = useServerFn(editDocument);

  // ── Auth ──────────────────────────────────────────────
  const {
    user,
    plan,
    filesCountToday,
    lastGenerationDate,
    loading: authLoading,
    signOut,
    incrementDailyFiles,
  } = useAuth();
  const navigate = useNavigate();

  // Redirect unauthenticated users to /login
  useEffect(() => {
    if (!authLoading && !user) {
      navigate({ to: "/login" });
    }
  }, [authLoading, user, navigate]);

  const [mode,   setMode]   = useState<DocMode>("word");
  const [prompt, setPrompt] = useState("");
  const [loading, setLoading] = useState(false);
  const [projects, setProjects] = useState<SavedProject[]>([]);
  const [menuOpen, setMenuOpen] = useState(false);

  const [activeProject, setActiveProject] = useState<SavedProject | null>(null);
  const [isManualEdit,  setIsManualEdit]   = useState(false);
  const [chatInput,     setChatInput]      = useState("");
  const [aiLoading,     setAiLoading]      = useState(false);
  const [activeSheetIdx, setActiveSheetIdx] = useState(0);
  const [activeSlideIdx, setActiveSlideIdx] = useState(0);

  useEffect(() => {
    document.documentElement.classList.add("dark");
    setProjects(getProjects());
  }, []);

  useEffect(() => {
    setActiveSheetIdx(0);
    setActiveSlideIdx(0);
  }, [activeProject?.id]);

  // Dynamic language check to format LTR / RTL globally
  const isArabic = activeProject ? isProjectArabic(activeProject) : false;

  // ── Auth gate: loading spinner ────────────────────────
  if (authLoading) {
    return (
      <div className="min-h-screen bg-background flex flex-col items-center justify-center gap-4">
        <div className="relative flex h-14 w-14 items-center justify-center rounded-2xl" style={{ background: "linear-gradient(135deg,#1e1e2e,#13131a)", border: "1px solid rgba(255,255,255,0.1)" }}>
          <Sparkles className="h-6 w-6 text-indigo-400 animate-pulse" />
        </div>
        <div className="flex items-center gap-2">
          <Loader2 className="h-4 w-4 animate-spin text-muted-foreground" />
          <span className="text-sm text-muted-foreground">جارٍ التحقق من الهوية…</span>
        </div>
      </div>
    );
  }

  // ── Auth gate: not logged in (redirect handled above) ─
  if (!user) return null;

  // ── Auth gate: free plan → Premium locked screen ───────
  if (plan !== "pro") {
    return (
      <div className="relative min-h-screen bg-background flex flex-col items-center justify-center px-4 overflow-hidden">
        {/* Decorative Grid */}
        <div className="pointer-events-none absolute inset-0 grid-bg opacity-30" />
        
        {/* Glow blobs */}
        <div className="pointer-events-none absolute inset-0 overflow-hidden">
          <div className="hero-glow-blob" style={{ width: "700px", height: "700px", top: "-250px", left: "-200px", background: "radial-gradient(circle, rgba(99,102,241,0.18) 0%, transparent 70%)" }} />
          <div className="hero-glow-blob" style={{ width: "500px", height: "500px", bottom: "-100px", right: "-100px", background: "radial-gradient(circle, rgba(249,115,22,0.12) 0%, transparent 70%)", animationDelay: "-3s" }} />
        </div>

        <motion.div
          initial={{ opacity: 0, y: 30, scale: 0.96 }}
          animate={{ opacity: 1, y: 0, scale: 1 }}
          transition={{ duration: 0.6, ease: [0.16, 1, 0.3, 1] }}
          className="relative z-10 w-full max-w-lg text-center"
        >
          {/* Locked Badge Icon with breathing glow */}
          <div className="mx-auto mb-8 relative flex h-24 w-24 items-center justify-center rounded-[2rem] border border-border"
               style={{
                 background: "linear-gradient(135deg, #181825 0%, #0d0d12 100%)",
                 boxShadow: "0 30px 60px -15px rgba(0,0,0,0.8), 0 0 40px rgba(99,102,241,0.15) inset"
               }}
          >
            <Lock className="h-9 w-9 text-indigo-400 animate-pulse" />
            <div className="absolute inset-0 rounded-[2rem] border-glow opacity-60" />
          </div>

          <h1 className="text-3xl sm:text-4xl font-black text-foreground dark:text-white tracking-tight leading-none mb-3">
            TOLZY Flow <span className="text-gradient font-extrabold">PRO</span>
          </h1>
          
          <p className="text-sm sm:text-base leading-relaxed text-[#8f8f9f] max-w-md mx-auto mb-8 direction-rtl" style={{ direction: "rtl" }}>
            هذه الأداة الاحترافية حصرية لمشتركي باقة <span className="text-foreground dark:text-white font-bold">TOLZY Pro</span> فقط.
            <br />
            <span className="text-xs text-muted-foreground mt-2 block">يرجى ترقية حسابك للحصول على وصول كامل وتجربة الأداة.</span>
          </p>

          {/* User Badge */}
          <div className="inline-flex items-center gap-2.5 rounded-full px-4 py-2 mb-8 border border-border/50"
               style={{ background: "rgba(255,255,255,0.02)" }}>
            <div className="h-2 w-2 rounded-full bg-indigo-500 animate-ping" />
            <span className="text-xs font-mono text-slate-700 dark:text-slate-300">{user.email}</span>
          </div>

          {/* Action CTAs */}
          <div className="flex flex-col sm:flex-row gap-4 justify-center items-center">
            <a
              href="https://tolzy.me/pricing"
              target="_blank"
              rel="noopener noreferrer"
              className="w-full sm:w-auto px-8 py-4 flex items-center justify-center gap-2.5 rounded-2xl text-sm font-bold text-[#050507] transition-all btn-shimmer border-glow"
              style={{
                background: "linear-gradient(135deg, #fff 0%, #e8e8ff 100%)",
                boxShadow: "0 10px 30px -10px rgba(255,255,255,0.25)"
              }}
            >
              <Sparkles className="h-4 w-4 text-indigo-600" />
              الترقية إلى TOLZY Pro
              <ArrowRight className="h-4 w-4 text-indigo-600 rotate-180" />
            </a>

            <button
              onClick={signOut}
              className="w-full sm:w-auto px-8 py-4 flex items-center justify-center gap-2 rounded-2xl text-sm font-semibold transition-all border border-border/50 text-[#8f8f9f] hover:text-foreground dark:text-white hover:bg-black/ dark:bg-white/[0.04]"
              style={{ background: "rgba(255,255,255,0.02)" }}
            >
              <LogOut className="h-4 w-4" />
              تسجيل الخروج
            </button>
          </div>
        </motion.div>
      </div>
    );
  }

  const handleMouseMove = (e: React.MouseEvent<HTMLElement>) => {
    const rect = e.currentTarget.getBoundingClientRect();
    e.currentTarget.style.setProperty("--mouse-x", `${e.clientX - rect.left}px`);
    e.currentTarget.style.setProperty("--mouse-y", `${e.clientY - rect.top}px`);
  };

  // ── Generate ──────────────────────────────────────────
  async function onGenerate() {
    if (plan !== "pro") {
      toast.error("يجب الاشتراك والترقية لتجربة الأداة.");
      return;
    }
    if (filesCountToday >= 9) {
      toast.error("لقد تجاوزت الحد اليومي الأقصى المسموح به لحسابك وهو 9 ملفات يومياً.");
      return;
    }
    if (prompt.trim().length < 8) { toast.error("أضف مزيدًا من التفاصيل."); return; }
    setLoading(true);
    try {
      const data = await generate({ data: { prompt, type: mode } });
      const id   = crypto.randomUUID();
      const proj =
        mode === "excel"
          ? saveProject({ id, title: data.title, type: "excel", sheets: data.sheets, prompt })
          : mode === "presentation"
          ? saveProject({ id, title: data.title, subtitle: data.subtitle, type: "presentation", slides: data.slides as FreestyleSlide[], prompt })
          : saveProject({ id, title: data.title, subtitle: data.subtitle, type: "word", sections: data.sections, prompt });
      
      // Persist generated file limit count
      await incrementDailyFiles();

      setProjects(getProjects());
      setActiveProject(proj);
      toast.success("تم التصميم بنجاح!");
    } catch (e: unknown) {
      toast.error(e instanceof Error ? e.message : "حدث خطأ");
    } finally {
      setLoading(false);
    }
  }

  // ── AI Edit ───────────────────────────────────────────
  async function onAIEdit() {
    if (plan !== "pro") {
      toast.error("يجب الاشتراك والترقية لتجربة الأداة.");
      return;
    }
    if (!activeProject || !chatInput.trim()) return;
    setAiLoading(true);
    const instruction = chatInput;
    setChatInput("");
    try {
      toast.info("الذكاء الاصطناعي يعالج طلبك…");
      const currentDoc =
        activeProject.type === "excel"
          ? { title: activeProject.title, sheets: activeProject.sheets }
          : activeProject.type === "presentation"
          ? { title: activeProject.title, subtitle: activeProject.subtitle, slides: activeProject.slides as FreestyleSlide[] }
          : { title: activeProject.title, subtitle: activeProject.subtitle, sections: activeProject.sections };
      const updated = await editWithAI({
        data: { currentDocument: currentDoc, instruction, type: activeProject.type },
      });
      const updatedProj =
        activeProject.type === "excel"
          ? saveProject({ ...activeProject, title: updated.title, sheets: updated.sheets })
          : activeProject.type === "presentation"
          ? saveProject({ ...activeProject, title: updated.title, subtitle: updated.subtitle, slides: updated.slides as FreestyleSlide[] })
          : saveProject({ ...activeProject, title: updated.title, subtitle: updated.subtitle, sections: updated.sections });
      setActiveProject(updatedProj);
      setProjects(getProjects());
      toast.success("تم تحديث المستند!");
    } catch (e: unknown) {
      toast.error(e instanceof Error ? e.message : "خطأ أثناء التعديل");
    } finally {
      setAiLoading(false);
    }
  }

  function handleManualSave() {
    if (!activeProject) return;
    saveProject(activeProject);
    setProjects(getProjects());
    toast.success("تم الحفظ محلياً.");
  }

  const updateProject = (fn: (p: SavedProject) => SavedProject) => {
    if (!activeProject) return;
    setActiveProject(fn(activeProject));
  };

  async function handleDownload() {
    if (plan !== "pro") {
      toast.error("يجب الاشتراك والترقية لتجربة الأداة.");
      return;
    }
    if (!activeProject) return;
    try {
      if (activeProject.type === "excel") {
        const name = `${activeProject.title.replace(/[^\w\- ]+/g, "").slice(0, 60) || "spreadsheet"}.xlsx`;
        await generateAndDownloadExcel({ title: activeProject.title, sheets: activeProject.sheets || [] }, name);
      } else if (activeProject.type === "presentation") {
        const name = `${activeProject.title.replace(/[^\w\- ]+/g, "").slice(0, 60) || "presentation"}.pptx`;
        await generateAndDownloadPptx(
          { title: activeProject.title, subtitle: activeProject.subtitle, slides: activeProject.slides || [] },
          name,
        );
      } else {
        const name = `${activeProject.title.replace(/[^\w\- ]+/g, "").slice(0, 60) || "document"}.docx`;
        await generateAndDownloadDocx(
          { title: activeProject.title, subtitle: activeProject.subtitle, sections: activeProject.sections || [] },
          name,
        );
      }
      toast.success("بدأ التحميل!");
    } catch {
      toast.error("فشل التحميل.");
    }
  }

  function handleDeleteProject(id: string, e: React.MouseEvent) {
    e.stopPropagation();
    setProjects(deleteProject(id));
    if (activeProject?.id === id) setActiveProject(null);
    toast.success("تم الحذف.");
  }

  // ── Excel helpers ─────────────────────────────────────
  function updateCell(sIdx: number, rIdx: number, cIdx: number, patch: Partial<ExcelCell>) {
    updateProject(prev => ({
      ...prev,
      sheets: prev.sheets!.map((s, si) =>
        si !== sIdx ? s : {
          ...s,
          rows: s.rows.map((row, ri) =>
            ri !== rIdx ? row : row.map((cell, ci) => (ci !== cIdx ? cell : { ...cell, ...patch }))
          ),
        }
      ),
    }));
  }

  function addExcelRow(sIdx: number) {
    updateProject(prev => ({
      ...prev,
      sheets: prev.sheets!.map((s, si) =>
        si !== sIdx ? s : { ...s, rows: [...s.rows, Array.from({ length: s.headers.length }, () => ({ value: "" }))] }
      ),
    }));
  }

  function deleteExcelRow(sIdx: number, rIdx: number) {
    updateProject(prev => ({
      ...prev,
      sheets: prev.sheets!.map((s, si) =>
        si !== sIdx ? s : { ...s, rows: s.rows.filter((_, ri) => ri !== rIdx) }
      ),
    }));
  }

  // Add Excel column in LTR/RTL order
  function addExcelColumn(sIdx: number) {
    updateProject(prev => ({
      ...prev,
      sheets: prev.sheets!.map((s, si) =>
        si !== sIdx ? s : {
          ...s,
          headers: [...s.headers, `العمود ${s.headers.length + 1}`],
          rows: s.rows.map(row => [...row, { value: "" }]),
        }
      ),
    }));
  }

  function deleteExcelColumn(sIdx: number, cIdx: number) {
    updateProject(prev => ({
      ...prev,
      sheets: prev.sheets!.map((s, si) =>
        si !== sIdx ? s : {
          ...s,
          headers: s.headers.filter((_, i) => i !== cIdx),
          rows: s.rows.map(row => row.filter((_, i) => i !== cIdx)),
        }
      ),
    }));
  }

  function addExcelSheet() {
    const newLen = (activeProject?.sheets?.length || 0) + 1;
    updateProject(prev => ({
      ...prev,
      sheets: [
        ...(prev.sheets || []),
        {
          name: `شيت ${newLen}`,
          headers: ["العمود الأول", "العمود الثاني", "العمود الثالث"],
          rows: [[{ value: "" }, { value: "" }, { value: "" }], [{ value: "" }, { value: "" }, { value: "" }]],
        },
      ],
    }));
    setTimeout(() => setActiveSheetIdx(newLen - 1), 50);
  }

  // ══════════════════════════════════════════════════════
  //  WORKSPACE VIEW
  // ══════════════════════════════════════════════════════
  if (activeProject) {
    const isExcel = activeProject.type === "excel";
    const isPpt   = activeProject.type === "presentation";
    const sheet   = isExcel ? (activeProject.sheets?.[activeSheetIdx] ?? null) : null;

    const accentColor = isExcel ? "#34d399" : isPpt ? "#fb923c" : "#818cf8";
    const accentBg    = isExcel ? "rgba(52,211,153,0.12)" : isPpt ? "rgba(251,146,60,0.12)" : "rgba(129,140,248,0.12)";
    const downloadBtnClass = isExcel
      ? "bg-emerald-600 hover:bg-emerald-500 shadow-emerald-900/30"
      : isPpt
      ? "bg-orange-600 hover:bg-orange-500 shadow-orange-900/30"
      : "bg-indigo-600 hover:bg-indigo-500 shadow-indigo-900/30";

    const previewAlignment = isArabic ? "text-right" : "text-left";
    const containerDirection = isArabic ? "rtl" : "ltr";

    // Sticky columns positioning for Excel preview grid
    const stickyColStyle: React.CSSProperties = {
      borderRight: "1px solid var(--excel-border)",
      borderLeft: "1px solid var(--excel-border)",
      borderBottom: "1px solid var(--excel-border)",
      background: "var(--excel-header-letter-bg)",
      position: "sticky",
      [isArabic ? "right" : "left"]: 0,
      zIndex: 30,
    };

    const rowNumStyle: React.CSSProperties = {
      borderRight: "1px solid var(--excel-border)",
      borderLeft: "1px solid var(--excel-border)",
      borderBottom: "1px solid var(--excel-border)",
      background: "var(--excel-header-letter-bg)",
      color: "var(--excel-header-letter-fg)",
      fontSize: "11px",
      textAlign: "center",
      fontFamily: "var(--font-mono)",
      position: "sticky",
      [isArabic ? "right" : "left"]: 0,
      zIndex: 10,
    };

    return (
      <main className="relative bg-background text-foreground flex flex-col overflow-hidden" style={{ height: "100dvh" }}>
        <Toaster theme="dark" position="top-center" richColors />

        {/* ── Workspace Toolbar ── */}
        <header className="shrink-0 h-14 nav-glass flex items-center justify-between px-4 gap-3 z-30">
          {/* Left: Back + Breadcrumb */}
          <div className="flex items-center gap-3 min-w-0">
            <button
              onClick={() => { saveProject(activeProject); setActiveProject(null); setProjects(getProjects()); }}
              className="flex items-center gap-1.5 text-xs text-muted-foreground hover:text-foreground dark:text-white glass px-3 py-1.5 rounded-xl transition-all hover:border-white/12 shrink-0"
            >
              <ChevronLeft className="h-3.5 w-3.5" />
              <span className="hidden sm:inline">الرئيسية</span>
            </button>

            <div className="flex items-center gap-1.5 text-muted-foreground shrink-0">
              <ChevronRight className="h-3 w-3" />
            </div>

            <div className="flex items-center gap-2.5 min-w-0 glass px-3 py-1.5 rounded-xl border border-white/6">
              <ModeIcon mode={activeProject.type} size="h-3.5 w-3.5" />
              <span className="text-xs font-semibold truncate max-w-[160px] text-foreground dark:text-white">{activeProject.title}</span>
              <TypeBadge type={activeProject.type} />
            </div>
          </div>

          {/* Right: Action buttons */}
          <div className="flex items-center gap-2 shrink-0">
            <button
              onClick={() => setIsManualEdit(v => !v)}
              className={`flex items-center gap-1.5 text-xs font-medium px-3 py-1.5 rounded-xl border transition-all ${
                isManualEdit
                  ? "bg-amber-500/15 border-amber-500/35 text-amber-300"
                  : "glass border-white/6 text-muted-foreground hover:text-foreground dark:text-white"
              }`}
            >
              {isManualEdit ? <Eye className="h-3.5 w-3.5" /> : <Edit className="h-3.5 w-3.5" />}
              <span className="hidden sm:inline">{isManualEdit ? "معاينة" : "تعديل يدوي"}</span>
            </button>
            {/* Sign out — workspace toolbar */}
            <button
              onClick={signOut}
              title="تسجيل الخروج"
              className="flex items-center gap-1.5 text-xs font-medium px-2.5 py-1.5 rounded-xl border transition-all glass border-white/6 text-muted-foreground hover:text-red-400 hover:border-red-500/20"
            >
              <LogOut className="h-3.5 w-3.5" />
              <span className="hidden sm:inline">خروج</span>
            </button>
            <button
              onClick={handleManualSave}
              className="flex items-center gap-1.5 text-xs font-semibold bg-black/ dark:bg-white/8 hover:bg-black/ dark:bg-white/12 text-foreground dark:text-white px-3 py-1.5 rounded-xl border border-white/8 transition-all"
            >
              <Save className="h-3.5 w-3.5" />
              <span className="hidden sm:inline">حفظ</span>
            </button>
            <button
              onClick={handleDownload}
              className={`flex items-center gap-1.5 text-xs font-semibold text-foreground dark:text-white px-4 py-1.5 rounded-xl transition-all shadow-lg btn-shimmer ${downloadBtnClass}`}
            >
              <Download className="h-3.5 w-3.5" />
              <span className="hidden sm:inline">تحميل {isExcel ? ".xlsx" : isPpt ? ".pptx" : ".docx"}</span>
            </button>
          </div>
        </header>

        {/* ── Body ── */}
        <div className="flex-1 flex flex-col md:flex-row overflow-hidden">

          {/* LEFT: AI Chat pane */}
          <div className="shrink-0 w-full md:w-[280px] sidebar-panel flex flex-col overflow-hidden">
            {/* Pane header */}
            <div className="px-4 py-3.5 border-b border-border/50 shrink-0 flex items-center gap-2">
              <div className="flex items-center justify-center w-7 h-7 rounded-lg" style={{ background: accentBg }}>
                <Cpu className="h-3.5 w-3.5" style={{ color: accentColor }} />
              </div>
              <div>
                <p className="text-xs font-bold text-foreground dark:text-white">لوحة الذكاء الاصطناعي</p>
                <p className="text-[9px] text-muted-foreground mt-0.5">AI Assistant Panel</p>
              </div>
            </div>

            {/* Scrollable content */}
            <div className="flex-1 overflow-y-auto p-4 space-y-4 custom-scrollbar">
              {/* Original prompt */}
              <div className="rounded-xl p-3.5" style={{ background: "rgba(255,255,255,0.025)", border: "1px solid rgba(255,255,255,0.055)" }}>
                <p className="text-[9px] uppercase tracking-widest font-bold mb-1.5" style={{ color: accentColor }}>
                  الوصف الأصلي
                </p>
                <p className="text-[11px] text-[#9898a8] italic leading-relaxed">
                  "{activeProject.prompt || "مستند تم إنشاؤه"}"
                </p>
              </div>

              {/* Quick tips label */}
              <div className="flex items-center gap-2">
                <div className="h-px flex-1" style={{ background: "rgba(255,255,255,0.05)" }} />
                <p className="text-[9px] font-bold uppercase tracking-widest text-muted-foreground">أمثلة سريعة</p>
                <div className="h-px flex-1" style={{ background: "rgba(255,255,255,0.05)" }} />
              </div>

              {/* Quick tip buttons */}
              <div className="space-y-1.5">
                {(isExcel
                  ? ["أضف عمود 'الضريبة 15%' بمعادلة", "أضف صف إجمالي بـ SUM", "أنشئ شيت ثانٍ ملخصاً", "نظم الأعمدة وأعد تسميتها"]
                  : isPpt
                  ? ["أضف شريحة للخاتمة والتواصل", "غير ألوان الشرائح لتكون درجات كحلي", "أضف شريحة تعرض المشكلة والحل", "اجعل النصوص أكثر إيجازاً ونقاطاً"]
                  : ["أضف قسم المخاطر المحتملة", "ترجم للعربية", "أضف جدولاً زمنياً", "اجعل الأسلوب أكثر رسمية"]
                ).map(tip => (
                  <button
                    key={tip}
                    onClick={() => setChatInput(tip)}
                    className="tip-bubble"
                    style={{ direction: "rtl" }}
                  >
                    {tip}
                  </button>
                ))}
              </div>

              {/* AI loading indicator */}
              {aiLoading && (
                <div className="flex items-center gap-3 p-3 rounded-xl" style={{ background: "rgba(129,140,248,0.08)", border: "1px solid rgba(129,140,248,0.2)" }}>
                  <Loader2 className="h-4 w-4 animate-spin text-indigo-400 shrink-0" />
                  <div>
                    <p className="text-[11px] font-semibold text-indigo-300">جارٍ التعديل…</p>
                    <p className="text-[9px] text-muted-foreground mt-0.5">الذكاء الاصطناعي يعالج طلبك</p>
                  </div>
                </div>
              )}
            </div>

            {/* Chat input */}
            <div className="shrink-0 p-4 border-t border-border/50 space-y-2.5" style={{ background: "rgba(0,0,0,0.4)" }}>
              <textarea
                value={chatInput}
                onChange={e => setChatInput(e.target.value)}
                onKeyDown={e => { if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) onAIEdit(); }}
                placeholder={
                  isExcel
                    ? "توجيهات لتحديث الجدول…"
                    : isPpt
                    ? "توجيهات لتحديث العرض التقديمي…"
                    : "توجيهات لتحديث المستند…"
                }
                rows={3}
                disabled={aiLoading}
                className="w-full resize-none rounded-xl text-xs leading-relaxed text-right text-foreground dark:text-white placeholder:text-muted-foreground outline-none transition-all"
                style={{
                  background: "rgba(255,255,255,0.03)",
                  border: "1px solid rgba(255,255,255,0.07)",
                  padding: "0.625rem 0.75rem",
                  direction: "rtl",
                  fontFamily: "var(--font-display)",
                }}
                onFocus={e => { e.target.style.borderColor = "rgba(255,255,255,0.14)"; e.target.style.background = "rgba(255,255,255,0.05)"; }}
                onBlur={e => { e.target.style.borderColor = "rgba(255,255,255,0.07)"; e.target.style.background = "rgba(255,255,255,0.03)"; }}
              />
              <button
                onClick={onAIEdit}
                disabled={aiLoading || !chatInput.trim()}
                className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-xs font-bold text-foreground dark:text-white transition-all disabled:opacity-35 disabled:cursor-not-allowed btn-shimmer"
                style={{
                  background: `linear-gradient(135deg, ${accentColor}30, ${accentColor}20)`,
                  border: `1px solid ${accentColor}40`,
                  color: accentColor,
                }}
              >
                <Sparkles className="h-3.5 w-3.5" />
                تعديل بالذكاء الاصطناعي
                <span className="opacity-40 text-[9px] font-normal">Ctrl+↵</span>
              </button>
            </div>
          </div>

          {/* RIGHT: Preview pane */}
          <div className="flex-1 overflow-hidden flex flex-col" style={{ background: "var(--excel-preview-bg)" }}>

            {/* ══ EXCEL ══ */}
            {isExcel && (
              <div className="flex-1 flex flex-col overflow-hidden">
                {/* Sheet tabs */}
                <div
                  className="shrink-0 flex items-center gap-0 border-b overflow-x-auto"
                  style={{ background: "var(--excel-sheet-tabs-bg)", borderColor: "var(--excel-border)" }}
                >
                  {(activeProject.sheets || []).map((s, si) => (
                    <button
                      key={si}
                      onClick={() => setActiveSheetIdx(si)}
                      className="flex items-center gap-1.5 px-5 py-2.5 text-xs whitespace-nowrap border-r transition-colors"
                      style={{
                        borderColor: "var(--excel-border)",
                        background:  si === activeSheetIdx ? "var(--excel-sheet-tab-active-bg)" : "transparent",
                        color:       si === activeSheetIdx ? "var(--excel-sheet-tab-active-fg)" : "var(--excel-sheet-tab-inactive-fg)",
                        borderBottom: si === activeSheetIdx ? "2px solid #34d399" : "2px solid transparent",
                      }}
                    >
                      <Table2 className="h-3 w-3 shrink-0" />
                      {isManualEdit ? (
                        <DebouncedInput
                          value={s.name}
                          onChange={val =>
                            updateProject(p => ({
                              ...p,
                              sheets: p.sheets!.map((sh, i) => (i === si ? { ...sh, name: val } : sh)),
                            }))
                          }
                          className="bg-transparent outline-none border-b w-20 text-center font-bold"
                          style={{ borderColor: "#34d399", color: "var(--excel-sheet-tab-active-fg)" }}
                          onClick={(e: any) => e.stopPropagation()}
                        />
                      ) : (
                        s.name
                      )}
                    </button>
                  ))}
                  {isManualEdit && (
                    <button
                      onClick={addExcelSheet}
                      className="flex items-center gap-1 px-4 py-2.5 text-xs whitespace-nowrap transition-colors"
                      style={{ color: "var(--excel-header-letter-fg)" }}
                      onMouseEnter={(e: any) => ((e.currentTarget as HTMLElement).style.color = "#34d399")}
                      onMouseLeave={(e: any) => ((e.currentTarget as HTMLElement).style.color = "var(--excel-header-letter-fg)")}
                    >
                      <Plus className="h-3.5 w-3.5" /> شيت جديد
                    </button>
                  )}
                </div>

                {/* Formula bar */}
                <div
                  className="shrink-0 flex items-center border-b"
                  style={{ background: "var(--excel-formula-bar-bg)", borderColor: "var(--excel-border)", height: "36px" }}
                >
                  <div
                    className="flex items-center justify-center px-4 border-r h-full gap-1"
                    style={{ borderColor: "var(--excel-border)", minWidth: "72px" }}
                  >
                    <span className="text-xs font-mono font-bold italic" style={{ color: "#34d399" }}>fx</span>
                  </div>
                  <div className="flex-1 px-4 text-[11px] font-mono flex items-center gap-2" style={{ color: "var(--excel-header-letter-fg)" }}>
                    {isManualEdit
                      ? "انقر على خلية لتحريرها · ابدأ بـ = لإدراج معادلة مثل: =SUM(B2:B10)"
                      : "وضع المعاينة · فعّل «تعديل يدوي» من الشريط العلوي لتحرير الخلايا"}
                  </div>
                  {isManualEdit && (
                    <div className="flex items-center gap-1 px-3 border-l h-full" style={{ borderColor: "var(--excel-border)" }}>
                      <button
                        onClick={() => addExcelColumn(activeSheetIdx)}
                        className="flex items-center gap-1 text-[11px] px-2 py-1 rounded transition"
                        style={{ color: "#34d399" }}
                        onMouseEnter={(e: any) => ((e.currentTarget as HTMLElement).style.background = "#064e3b44")}
                        onMouseLeave={(e: any) => ((e.currentTarget as HTMLElement).style.background = "transparent")}
                      >
                        <Plus className="h-3 w-3" /> عمود
                      </button>
                      <button
                        onClick={() => addExcelRow(activeSheetIdx)}
                        className="flex items-center gap-1 text-[11px] px-2 py-1 rounded transition"
                        style={{ color: "#34d399" }}
                        onMouseEnter={(e: any) => ((e.currentTarget as HTMLElement).style.background = "#064e3b44")}
                        onMouseLeave={(e: any) => ((e.currentTarget as HTMLElement).style.background = "transparent")}
                      >
                        <Plus className="h-3 w-3" /> صف
                      </button>
                    </div>
                  )}
                </div>

                {/* Grid */}
                <div className="flex-1 overflow-auto">
                  {sheet && (
                    <table
                      className="border-collapse"
                      dir={containerDirection}
                      style={{
                        tableLayout: "fixed",
                        width: "100%",
                        minWidth: `${sheet.headers.length * 160 + 52}px`,
                      }}
                    >
                      <colgroup>
                        <col style={{ width: "52px" }} />
                        {sheet.headers.map((_, i) => <col key={i} style={{ minWidth: "130px" }} />)}
                        {isManualEdit && <col style={{ width: "36px" }} />}
                      </colgroup>

                      <thead style={{ position: "sticky", top: 0, zIndex: 20 }}>
                        {/* Column letter row */}
                        <tr style={{ background: "var(--excel-header-letter-bg)", height: "20px" }}>
                          <th style={stickyColStyle} />
                          {sheet.headers.map((_, hi) => (
                            <th
                              key={hi}
                              style={{
                                borderRight: "1px solid var(--excel-border)",
                                borderBottom: "1px solid var(--excel-border)",
                                color: "var(--excel-header-letter-fg)",
                                fontSize: "10px",
                                fontWeight: 400,
                                textAlign: "center",
                                fontFamily: "var(--font-mono)",
                              }}
                            >
                              {colLabel(hi)}
                            </th>
                          ))}
                          {isManualEdit && <th style={{ borderBottom: "1px solid var(--excel-border)", background: "var(--excel-header-letter-bg)" }} />}
                        </tr>

                        {/* Header / column name row */}
                        <tr style={{ height: "44px" }}>
                          <td style={stickyColStyle} />
                          {sheet.headers.map((header, hi) => (
                            <th
                              key={hi}
                              className="group/hdr relative"
                              style={{
                                borderRight: "1px solid var(--excel-border)",
                                borderBottom: "2px solid #166534",
                                background: "linear-gradient(180deg, #162a1e, #0e1f15)",
                                padding: 0,
                              }}
                            >
                              <div className="flex items-center justify-between h-full px-3 gap-1" style={{ height: "44px" }}>
                                {isManualEdit ? (
                                  <DebouncedInput
                                    value={header}
                                    onChange={val =>
                                      updateProject(p => ({
                                        ...p,
                                        sheets: p.sheets!.map((s, si) =>
                                          si !== activeSheetIdx
                                            ? s
                                            : { ...s, headers: s.headers.map((h, i) => (i === hi ? val : h)) }
                                        ),
                                      }))
                                    }
                                    className="flex-1 bg-transparent outline-none text-center font-semibold min-w-0"
                                    style={{ color: "#86efac", caretColor: "#34d399", fontSize: "12px", textAlign: isArabic ? "right" : "left" }}
                                  />
                                ) : (
                                  <span className={`flex-1 font-semibold truncate ${previewAlignment}`} style={{ color: "#86efac", fontSize: "12px" }}>
                                    {header}
                                  </span>
                                )}
                                {isManualEdit && (
                                  <button
                                    onClick={() => deleteExcelColumn(activeSheetIdx, hi)}
                                    className="opacity-0 group-hover/hdr:opacity-100 transition shrink-0 rounded p-0.5"
                                    style={{ color: "#f87171" }}
                                    onMouseEnter={(e: any) => ((e.currentTarget as HTMLElement).style.background = "#450a0a44")}
                                    onMouseLeave={(e: any) => ((e.currentTarget as HTMLElement).style.background = "transparent")}
                                  >
                                    <Trash2 className="h-2.5 w-2.5" />
                                  </button>
                                )}
                              </div>
                            </th>
                          ))}
                          {isManualEdit && <td style={{ borderBottom: "2px solid #166534", background: "#0e1f15" }} />}
                        </tr>
                      </thead>

                      <tbody>
                        {sheet.rows.map((row, ri) => {
                          const even = ri % 2 === 0;
                          const bg   = even ? "var(--excel-row-even)" : "var(--excel-row-odd)";
                          return (
                            <tr
                              key={ri}
                              className="group/row"
                              style={{ background: bg, height: "34px" }}
                              onMouseEnter={(e: any) => ((e.currentTarget as HTMLElement).style.background = "var(--excel-row-hover)")}
                              onMouseLeave={(e: any) => ((e.currentTarget as HTMLElement).style.background = bg)}
                            >
                              {/* Row number */}
                              <td style={rowNumStyle}>
                                {ri + 1}
                              </td>

                              {Array.from({ length: sheet.headers.length }).map((_, ci) => {
                                const cell: ExcelCell = row[ci] ?? { value: "" };
                                const isFormula = !!cell.formula;
                                const display   = isFormula ? cell.formula : cell.value;
                                const bg2       = cell.bg ? `#${cell.bg.replace("#", "")}28` : "transparent";
                                
                                // Determine text color and alignments based on language/content
                                const isNumeric = 
                                  typeof cell.value === "number" || 
                                  (!isNaN(Number(cell.value)) && String(cell.value).trim() !== "" && !cell.value.toString().startsWith("0"));

                                const fgColor   = cell.color
                                  ? `#${cell.color.replace("#", "")}`
                                  : isFormula
                                  ? "#10b981"
                                  : "var(--foreground)";
                                return (
                                  <td
                                    key={ci}
                                    style={{
                                      borderRight: "1px solid var(--excel-border)",
                                      borderBottom: "1px solid var(--excel-border)",
                                      background: bg2,
                                      padding: 0,
                                      height: "34px",
                                    }}
                                  >
                                    {isManualEdit ? (
                                      <DebouncedInput
                                        value={cell.formula || String(cell.value ?? "")}
                                        onChange={val => {
                                          val.startsWith("=")
                                            ? updateCell(activeSheetIdx, ri, ci, { formula: val, value: "" })
                                            : updateCell(activeSheetIdx, ri, ci, { value: val, formula: undefined });
                                        }}
                                        className="w-full h-full bg-transparent outline-none px-3"
                                        style={{
                                          color: fgColor,
                                          fontWeight: cell.bold ? 700 : 400,
                                          fontSize: "12px",
                                          caretColor: "#34d399",
                                          textAlign: isNumeric ? "right" : (isArabic ? "right" : "left"),
                                        }}
                                        onFocus={(e: any) => {
                                          const td = e.target.closest("td") as HTMLElement;
                                          if (td) td.style.boxShadow = "inset 0 0 0 2px #34d399";
                                        }}
                                        onBlur={(e: any) => {
                                          const td = e.target.closest("td") as HTMLElement;
                                          if (td) td.style.boxShadow = "none";
                                        }}
                                      />
                                    ) : (
                                      <div
                                        className={`w-full h-full flex items-center px-3 overflow-hidden ${
                                          isNumeric ? "justify-end" : (isArabic ? "justify-start" : "justify-start")
                                        }`}
                                        title={String(display ?? "")}
                                        style={{
                                          color: fgColor,
                                          fontWeight: cell.bold ? 700 : 400,
                                          fontSize: "12px",
                                          textAlign: isNumeric ? "right" : (isArabic ? "right" : "left"),
                                        }}
                                      >
                                        <span className="truncate">{String(display ?? "")}</span>
                                      </div>
                                    )}
                                  </td>
                                );
                              })}

                              {isManualEdit && (
                                <td style={{ borderBottom: "1px solid var(--excel-border)", background: "transparent", textAlign: "center" }}>
                                  <button
                                    onClick={() => deleteExcelRow(activeSheetIdx, ri)}
                                    className="opacity-0 group-hover/row:opacity-100 transition p-1 rounded"
                                    style={{ color: "#ef4444" }}
                                    onMouseEnter={(e: any) => ((e.currentTarget as HTMLElement).style.background = "#450a0a44")}
                                    onMouseLeave={(e: any) => ((e.currentTarget as HTMLElement).style.background = "transparent")}
                                  >
                                    <Trash2 className="h-3 w-3" />
                                  </button>
                                </td>
                              )}
                            </tr>
                          );
                        })}

                        {isManualEdit && (
                          <tr style={{ background: "var(--excel-sheet-tabs-bg)", height: "34px" }}>
                            <td style={{ borderRight: "1px solid var(--excel-border)", borderTop: "1px dashed var(--excel-border)", background: "var(--excel-row-even)", position: "sticky", [isArabic ? "right" : "left"]: 0 }} />
                            <td colSpan={sheet.headers.length + 1} style={{ borderTop: "1px dashed var(--excel-border)" }}>
                              <button
                                onClick={() => addExcelRow(activeSheetIdx)}
                                className="w-full flex items-center gap-2 px-4 py-2 text-xs transition-colors"
                                style={{ color: "var(--excel-header-letter-fg)" }}
                                onMouseEnter={(e: any) => {
                                  (e.currentTarget as HTMLElement).style.color = "#34d399";
                                  (e.currentTarget as HTMLElement).style.background = "#064e3b22";
                                }}
                                onMouseLeave={(e: any) => {
                                  (e.currentTarget as HTMLElement).style.color = "var(--excel-header-letter-fg)";
                                  (e.currentTarget as HTMLElement).style.background = "transparent";
                                }}
                              >
                                <PlusCircle className="h-3.5 w-3.5" /> إضافة صف جديد
                              </button>
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>
            )}

            {/* ══ POWERPOINT ══ */}
            {isPpt && (
              <div className="flex-1 flex flex-col overflow-hidden" style={{ background: "#070709" }}>
                {/* Main Slide canvas */}
                <div className="flex-1 p-6 flex items-center justify-center overflow-y-auto">
                  {(() => {
                    const slides      = (activeProject.slides || []) as FreestyleSlide[];
                    const activeSlide = slides[activeSlideIdx] || null;
                    if (!activeSlide) return null;

                    const accentColor = activeSlide.accentColor || "#6366F1";
                    const bgUrl = activeSlide.bgImagePath?.startsWith("http")
                      ? activeSlide.bgImagePath
                      : undefined;

                    return (
                      <motion.div
                        key={activeSlideIdx}
                        initial={{ opacity: 0, scale: 0.97 }}
                        animate={{ opacity: 1, scale: 1 }}
                        transition={{ duration: 0.28, ease: "easeOut" }}
                        className="relative w-full max-w-[840px] aspect-[16/9] rounded-2xl shadow-2xl overflow-hidden border border-white/6"
                        dir={containerDirection}
                        style={{ boxShadow: "0 40px 120px -30px rgba(0,0,0,0.95), 0 0 0 1px rgba(255,255,255,0.04) inset" }}
                      >
                        {/* Background image */}
                        {bgUrl ? (
                          <img
                            src={bgUrl}
                            alt=""
                            className="absolute inset-0 w-full h-full object-cover"
                          />
                        ) : (
                          <div className="absolute inset-0" style={{ background: `linear-gradient(135deg, #0f172a, #1e293b)` }} />
                        )}

                        {/* Overlay */}
                        <div className="absolute inset-0" style={{ background: "rgba(0,0,0,0.52)" }} />

                        {/* Accent top bar */}
                        <div className="absolute top-0 left-0 right-0 h-1" style={{ background: accentColor }} />

                        {/* Elements preview */}
                        <div className="absolute inset-0 p-4">
                          {/* Canvas ratio helper: 840px wide = 13.33in, so 1in ≈ 63px */}
                          {(activeSlide.elements || []).map((el, ei) => {
                            const SCALE = 840 / 13.33; // px per inch in preview
                            const H_SCALE = (840 * 9/16) / 7.5;
                            const style: React.CSSProperties = {
                              position: "absolute",
                              left: `${el.x * SCALE}px`,
                              top: `${el.y * H_SCALE}px`,
                              width: `${el.w * SCALE}px`,
                              height: `${el.h * H_SCALE}px`,
                              overflow: "hidden",
                            };

                            if (el.type === "hero_text") {
                              return (
                                <div key={ei} style={{ ...style, display: "flex", alignItems: "center" }}>
                                  <span style={{
                                    fontSize: `${(el.fontSize ?? 40) * (SCALE / 96)}px`,
                                    fontWeight: 800,
                                    color: el.color || "#FFFFFF",
                                    lineHeight: 1.15,
                                    textShadow: "0 4px 12px rgba(0,0,0,0.7)",
                                    direction: isArabic ? "rtl" : "ltr",
                                  }}>
                                    {el.text}
                                  </span>
                                </div>
                              );
                            }

                            if (el.type === "subtext") {
                              return (
                                <div key={ei} style={{ ...style, display: "flex", alignItems: "center" }}>
                                  <span style={{
                                    fontSize: `${(el.fontSize ?? 14) * (SCALE / 96)}px`,
                                    color: el.color || "#94A3B8",
                                    lineHeight: 1.55,
                                    direction: isArabic ? "rtl" : "ltr",
                                  }}>
                                    {el.text}
                                  </span>
                                </div>
                              );
                            }

                            if (el.type === "glass_card") {
                              return (
                                <div key={ei} style={{
                                  ...style,
                                  background: "rgba(13,17,23,0.72)",
                                  backdropFilter: "blur(12px)",
                                  border: `1px solid ${accentColor}55`,
                                  borderRadius: "10px",
                                  borderTop: `2px solid ${accentColor}`,
                                  padding: "8px 10px",
                                  display: "flex",
                                  flexDirection: "column",
                                  gap: "4px",
                                  direction: isArabic ? "rtl" : "ltr",
                                }}>
                                  {el.title && (
                                    <div style={{ fontSize: "9px", fontWeight: 700, color: accentColor, marginBottom: "2px" }}>
                                      {el.title}
                                    </div>
                                  )}
                                  <div style={{ fontSize: "8px", color: "#CBD5E1", lineHeight: 1.5, overflow: "hidden" }}>
                                    {el.text}
                                  </div>
                                </div>
                              );
                            }

                            if (el.type === "image_node") {
                              const imgUrl = el.imageUrl?.startsWith("http") ? el.imageUrl : getUnsplashImageUrl(el.imageUrl);
                              return (
                                <div key={ei} style={{ ...style, borderRadius: "8px", overflow: "hidden", border: `1px solid ${accentColor}40` }}>
                                  <img src={imgUrl} alt="" className="w-full h-full object-cover" />
                                </div>
                              );
                            }

                            if (el.type === "kpi_block") {
                              return (
                                <div key={ei} style={{
                                  ...style,
                                  background: "rgba(13,17,23,0.65)",
                                  border: `1px solid ${accentColor}44`,
                                  borderLeft: `3px solid ${accentColor}`,
                                  borderRadius: "8px",
                                  padding: "6px 10px",
                                  display: "flex",
                                  flexDirection: "column",
                                  justifyContent: "center",
                                  direction: isArabic ? "rtl" : "ltr",
                                }}>
                                  <div style={{ fontSize: `${Math.max(14, Math.min(28, 60 / Math.max(el.value.length, 1))) * (SCALE / 96)}px`, fontWeight: 800, color: accentColor, lineHeight: 1 }}>
                                    {el.value}
                                  </div>
                                  <div style={{ fontSize: "7px", color: "#94A3B8", marginTop: "3px" }}>
                                    {el.label}
                                  </div>
                                </div>
                              );
                            }

                            if (el.type === "accent_shape") {
                              const shapeColor = el.color || accentColor;
                              const opacity = el.opacity ?? 0.15;
                              return (
                                <div key={ei} style={{
                                  ...style,
                                  background: shapeColor,
                                  opacity,
                                  borderRadius: el.shape === "rect" ? "4px" : "50%",
                                }} />
                              );
                            }

                            if (el.type === "divider_line") {
                              return (
                                <div key={ei} style={{
                                  ...style,
                                  background: el.color || "#334155",
                                  borderRadius: "2px",
                                }} />
                              );
                            }

                            return null;
                          })}
                        </div>

                        {/* Slide number + branding */}
                        <div className="absolute bottom-2 left-4 right-4 flex justify-between items-center" style={{ zIndex: 10 }}>
                          <span style={{
                            fontSize: "9px",
                            background: `${accentColor}33`,
                            border: `1px solid ${accentColor}55`,
                            color: "var(--foreground)",
                            padding: "2px 7px",
                            borderRadius: "4px",
                            fontWeight: 700,
                          }}>
                            {String(activeSlideIdx + 1).padStart(2, "0")} / {String(slides.length).padStart(2, "0")}
                          </span>
                          <span style={{ fontSize: "8px", color: "rgba(255,255,255,0.35)", fontWeight: 600 }}>TOLZY Flow</span>
                        </div>
                      </motion.div>
                    );
                  })()}
                </div>

                {/* Slides deck thumbnail strip */}
                <div
                  className="shrink-0 h-[108px] border-t border-border/50 flex items-center px-5 gap-3 overflow-x-auto custom-scrollbar"
                  style={{ background: "rgba(0,0,0,0.5)", backdropFilter: "blur(20px)" }}
                >
                  {((activeProject.slides || []) as FreestyleSlide[]).map((slide, si) => {
                    const acc = slide.accentColor || "#6366F1";
                    const bgImg = slide.bgImagePath?.startsWith("http") ? slide.bgImagePath : undefined;
                    return (
                      <button
                        key={si}
                        onClick={() => setActiveSlideIdx(si)}
                        className={`slide-thumb ${si === activeSlideIdx ? "slide-thumb-active" : ""}`}
                        title={si === 0 ? activeProject.title : slide.slide_title}
                      >
                        <div
                          className="w-full h-full flex flex-col justify-center items-center p-2 relative overflow-hidden"
                          style={si === 0
                            ? { backgroundColor: "#0F172A" }
                            : { backgroundColor: "#111827" }
                          }
                        >
                          {bgImg && (
                            <img src={bgImg} alt="" className="absolute inset-0 w-full h-full object-cover opacity-40" />
                          )}
                          <div className="absolute inset-x-0 top-0 h-0.5" style={{ background: acc }} />
                          <p className="text-[8px] text-foreground dark:text-white font-bold leading-tight text-center line-clamp-2 relative z-10">
                            {si === 0 ? activeProject.title : slide.slide_title}
                          </p>
                          <p className="text-[6px] text-white/40 mt-1 relative z-10">{si + 1}</p>
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>
            )}

            {/* ══ WORD ══ */}
            {!isExcel && !isPpt && (
              <div className="flex-1 overflow-y-auto p-6 sm:p-10 flex justify-center custom-scrollbar" style={{ background: "var(--word-preview-bg)" }}>
                <div
                  className={`relative w-full max-w-[820px] min-h-[1050px] word-page p-10 sm:p-16 ${previewAlignment}`}
                  dir={containerDirection}
                >

                  {isManualEdit && (
                    <div className="absolute top-4 left-4 right-4 bg-amber-500/8 border border-amber-500/25 text-amber-300 text-[11px] px-4 py-2.5 rounded-xl flex items-center justify-between z-10">
                      <span>💡 <strong>وضع التعديل اليدوي:</strong> انقر مباشرة على أي نص لتعديله.</span>
                      <button onClick={() => setIsManualEdit(false)} className="font-bold hover:text-amber-100 ml-3">✕</button>
                    </div>
                  )}

                  {/* Document title area */}
                  <div className="text-center mb-12 mt-4">
                    {isManualEdit ? (
                      <DebouncedInput
                        value={activeProject.title}
                        onChange={val => updateProject(p => ({ ...p, title: val }))}
                        className="w-full text-center text-3xl sm:text-4xl font-black outline-none pb-1 bg-transparent border-b-2 border-indigo-500/30 focus:border-indigo-400 transition-colors"
                        style={{ color: "#818cf8" }}
                      />
                    ) : (
                      <h1 className="text-3xl sm:text-4xl font-black leading-tight" style={{ color: "#818cf8" }}>
                        {activeProject.title}
                      </h1>
                    )}
                    <div className="mt-2 mx-auto h-0.5 w-16 rounded-full" style={{ background: "linear-gradient(90deg,#6366f1,#818cf8)" }} />
                    <div className="mt-4">
                      {isManualEdit ? (
                        <DebouncedInput
                          placeholder="العنوان الفرعي (اختياري)…"
                          value={activeProject.subtitle || ""}
                          onChange={val => updateProject(p => ({ ...p, subtitle: val }))}
                          className="w-full text-center text-sm outline-none bg-transparent border-b border-white/8 focus:border-black/20 dark:border-white/20 transition-colors italic"
                          style={{ color: "#6b6b7b" }}
                        />
                      ) : activeProject.subtitle ? (
                        <p className="text-sm italic" style={{ color: "#6b6b7b" }}>{activeProject.subtitle}</p>
                      ) : null}
                    </div>
                  </div>

                  {/* Sections */}
                  <div className="space-y-10">
                    {(activeProject.sections || []).map((sec, si) => (
                      <div
                        key={si}
                        className="group relative"
                      >
                        {/* Section side accent bar */}
                        <div className={`absolute top-0 bottom-0 w-0.5 rounded-full opacity-0 group-hover:opacity-100 transition-opacity duration-300 ${
                          isArabic ? "-right-8" : "-left-8"
                        }`} style={{ background: "linear-gradient(180deg, #6366f1, transparent)" }} />

                        {isManualEdit && (
                          <div className={`absolute -top-3 flex items-center gap-1 opacity-0 group-hover:opacity-100 glass border border-border shadow-xl rounded-lg p-1 z-10 transition ${
                            isArabic ? "-left-1" : "-right-1"
                          }`}>
                            <button
                              onClick={() => updateProject(p => { const s = [...p.sections!]; s.splice(si, 1); return { ...p, sections: s }; })}
                              className="text-red-400 hover:bg-red-500/10 p-1 rounded"
                              title="حذف القسم"
                            >
                              <Trash2 className="h-3.5 w-3.5" />
                            </button>
                            <button
                              onClick={() =>
                                updateProject(p => {
                                  const s = [...p.sections!];
                                  s[si] = { ...s[si], paragraphs: [...(s[si].paragraphs || []), "فقرة جديدة…"] };
                                  return { ...p, sections: s };
                                })
                              }
                              className="text-[#9898a8] hover:bg-black/ dark:bg-black/5 dark:bg-white/5 p-1 rounded text-[10px] flex items-center gap-0.5"
                            >
                              <Plus className="h-3 w-3" />فقرة
                            </button>
                          </div>
                        )}

                        {isManualEdit ? (
                          <DebouncedInput
                            value={sec.heading}
                            onChange={val =>
                              updateProject(p => {
                                const s = [...p.sections!];
                                s[si] = { ...s[si], heading: val };
                                return { ...p, sections: s };
                              })
                            }
                            className={`w-full text-xl font-bold outline-none pb-1 bg-transparent border-b border-white/8 focus:border-indigo-400/50 transition-colors mb-3 font-bold ${previewAlignment}`}
                            style={{ color: "#a5b4fc" }}
                          />
                        ) : (
                          <h2 className="text-xl font-bold mb-3" style={{ color: "#a5b4fc" }}>{sec.heading}</h2>
                        )}

                        {isManualEdit && (
                          <DebouncedInput
                            placeholder="عنوان فرعي (اختياري)…"
                            value={sec.subheading || ""}
                            onChange={val =>
                              updateProject(p => {
                                const s = [...p.sections!];
                                s[si] = { ...s[si], subheading: val };
                                return { ...p, sections: s };
                              })
                            }
                            className={`w-full text-sm font-semibold outline-none bg-transparent border-b border-white/6 focus:border-white/18 transition-colors mb-3 ${previewAlignment}`}
                            style={{ color: "#52526a" }}
                          />
                        )}
                        {!isManualEdit && sec.subheading && (
                          <h3 className="text-sm font-semibold uppercase tracking-widest mb-3" style={{ color: "#52526a" }}>
                            {sec.subheading}
                          </h3>
                        )}

                        <div className="space-y-3 mt-1">
                          {(sec.paragraphs ?? []).map((para, pi) => (
                            <div key={pi} className={`flex gap-2 items-start ${previewAlignment}`} style={{ direction: containerDirection }}>
                              {isManualEdit ? (
                                <>
                                  <DebouncedTextarea
                                    value={para}
                                    rows={Math.max(2, Math.ceil(para.length / 90))}
                                    onChange={val =>
                                      updateProject(p => {
                                        const s = [...p.sections!];
                                        const paragraphs = [...(s[si].paragraphs || [])];
                                        paragraphs[pi] = val;
                                        s[si] = { ...s[si], paragraphs };
                                        return { ...p, sections: s };
                                      })
                                    }
                                    className={`flex-1 text-sm p-2.5 rounded-lg border outline-none resize-none transition-all ${previewAlignment}`}
                                    style={{
                                      color: "var(--muted-foreground)",
                                      background: "rgba(255,255,255,0.025)",
                                      borderColor: "rgba(255,255,255,0.08)",
                                    }}
                                    onFocus={(e: any) => { e.target.style.borderColor = "rgba(129,140,248,0.4)"; }}
                                    onBlur={(e: any) => { e.target.style.borderColor = "rgba(255,255,255,0.08)"; }}
                                  />
                                  <button
                                    onClick={() =>
                                      updateProject(p => {
                                        const s = [...p.sections!];
                                        const paragraphs = (s[si].paragraphs || []).filter((_, i) => i !== pi);
                                        s[si] = { ...s[si], paragraphs };
                                        return { ...p, sections: s };
                                      })
                                    }
                                    className="p-1.5 mt-1"
                                    style={{ color: "#ef4444" }}
                                  >
                                    <MinusCircle className="h-4 w-4" />
                                  </button>
                                </>
                              ) : (
                                <p className="leading-loose text-[15px] text-justify flex-1" style={{ color: "var(--muted-foreground)" }}>
                                  {para}
                                </p>
                              )}
                            </div>
                          ))}
                        </div>

                        {sec.table && sec.table.headers?.length > 0 && (
                          <div className="mt-6 overflow-x-auto border border-white/8 rounded-xl">
                            <table className="w-full border-collapse text-xs" dir={containerDirection}>
                              <thead>
                                <tr className="border-b border-white/8" style={{ background: "rgba(99,102,241,0.1)" }}>
                                  {sec.table.headers.map((h, hi) => (
                                    <th key={hi} className={`p-3 font-semibold border border-white/8 ${previewAlignment}`} style={{ color: "#a5b4fc" }}>
                                      {isManualEdit ? (
                                        <DebouncedInput
                                          value={h}
                                          onChange={val =>
                                            updateProject(p => {
                                              const s = [...p.sections!];
                                              const tbl = { ...s[si].table! };
                                              const headers = [...tbl.headers];
                                              headers[hi] = val;
                                              s[si] = { ...s[si], table: { ...tbl, headers } };
                                              return { ...p, sections: s };
                                            })
                                          }
                                          className={`w-full bg-transparent border-b outline-none font-semibold ${previewAlignment}`}
                                          style={{ borderColor: "rgba(129,140,248,0.3)", color: "var(--foreground)" }}
                                        />
                                      ) : (
                                        h
                                      )}
                                    </th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody>
                                {sec.table.rows.map((row, ri) => (
                                  <tr
                                    key={ri}
                                    style={{ background: ri % 2 === 1 ? "rgba(255,255,255,0.02)" : "transparent" }}
                                  >
                                    {Array.from({ length: sec.table!.headers.length }).map((_, ci) => (
                                      <td key={ci} className={`p-3 border border-white/8 ${previewAlignment}`} style={{ color: "#9898a8" }}>
                                        {isManualEdit ? (
                                          <DebouncedInput
                                            value={row[ci] ?? ""}
                                            onChange={val =>
                                              updateProject(p => {
                                                const s = [...p.sections!];
                                                const tbl = { ...s[si].table! };
                                                const rows = tbl.rows.map((r, rj) =>
                                                  rj === ri ? r.map((c, cj) => (cj === ci ? val : c)) : r
                                                );
                                                s[si] = { ...s[si], table: { ...tbl, rows } };
                                                return { ...p, sections: s };
                                              })
                                            }
                                            className={`w-full bg-transparent outline-none border-b border-transparent focus:border-indigo-500/30 ${previewAlignment}`}
                                            style={{ color: "var(--muted-foreground)" }}
                                          />
                                        ) : (
                                          row[ci] ?? ""
                                        )}
                                      </td>
                                    ))}
                                    {isManualEdit && (
                                      <td className="border border-white/8 text-center w-8">
                                        <button
                                          onClick={() =>
                                            updateProject(p => {
                                              const s = [...p.sections!];
                                              const tbl = { ...s[si].table! };
                                              const rows = tbl.rows.filter((_, rj) => rj !== ri);
                                              s[si] = { ...s[si], table: { ...tbl, rows } };
                                              return { ...p, sections: s };
                                            })
                                          }
                                          className="p-1"
                                          style={{ color: "#ef4444" }}
                                        >
                                          <Trash2 className="h-3 w-3" />
                                        </button>
                                      </td>
                                    )}
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}

          </div>
        </div>
      </main>
    );
  }

  // ══════════════════════════════════════════════════════════════
  //  LANDING PAGE
  // ══════════════════════════════════════════════════════════════
  const modeConfig = {
    word:         { label: "Document",     color: "#818cf8", icon: <FileText className="h-4 w-4" />, example: WORD_EXAMPLES },
    excel:        { label: "Spreadsheet",  color: "#34d399", icon: <FileSpreadsheet className="h-4 w-4" />, example: EXCEL_EXAMPLES },
    presentation: { label: "Presentation", color: "#fb923c", icon: <Presentation className="h-4 w-4" />, example: PRESENTATION_EXAMPLES },
  };
  const currentMode = modeConfig[mode];

  return (
    <main className="relative min-h-screen text-foreground bg-background overflow-x-hidden selection:bg-indigo-500/25">
      <Toaster theme="dark" position="top-center" richColors />

      {/* Dot grid background */}
      <div className="pointer-events-none absolute inset-0 grid-bg opacity-50" />

      {/* Ambient glow blobs */}
      <div className="pointer-events-none absolute inset-0 overflow-hidden">
        <div
          className="hero-glow-blob"
          style={{
            width: "700px", height: "700px",
            top: "-200px", left: "-150px",
            background: "radial-gradient(circle, rgba(99,102,241,0.14) 0%, transparent 70%)",
            animationDelay: "0s",
          }}
        />
        <div
          className="hero-glow-blob"
          style={{
            width: "500px", height: "500px",
            top: "20%", right: "-100px",
            background: "radial-gradient(circle, rgba(249,115,22,0.10) 0%, transparent 70%)",
            animationDelay: "-4s",
          }}
        />
        <div
          className="hero-glow-blob"
          style={{
            width: "400px", height: "400px",
            bottom: "5%", left: "30%",
            background: "radial-gradient(circle, rgba(16,185,129,0.08) 0%, transparent 70%)",
            animationDelay: "-2s",
          }}
        />
      </div>

      {/* ── Navbar ── */}
      <nav className="fixed top-0 inset-x-0 h-16 nav-glass z-50 flex items-center justify-between px-6">
        {/* Logo */}
        <div className="flex items-center gap-3">
          <div className="relative flex h-10 w-10 items-center justify-center rounded-xl border border-border overflow-hidden transition-all duration-300 hover:scale-105 hover:shadow-[0_0_15px_rgba(99,102,241,0.2)] bg-[#0A0A0F]">
            <img src="/image/logo.jpg" alt="TOLZY Flow Logo" className="w-full h-full object-cover" />
            <div className="absolute inset-0 rounded-xl border border-border/50 pointer-events-none" />
          </div>
          <div>
            <p className="text-sm font-black tracking-widest font-mono uppercase bg-clip-text text-transparent bg-gradient-to-r from-foreground to-foreground/70 dark:from-white dark:to-white/70">TOLZY Flow</p>
            <p className="text-[9px] text-indigo-400/80 leading-none font-medium tracking-wider mt-0.5"></p>
          </div>
        </div>

        {/* Right nav */}
        <div className="flex items-center gap-3">
          
          {/* Theme Toggle */}
          <button
            onClick={() => setIsDark(!isDark)}
            className="flex h-9 w-9 items-center justify-center rounded-full transition-all border border-black/5 dark:border-border text-slate-500 hover:text-slate-900 dark:text-slate-700 dark:text-slate-300 dark:hover:text-foreground dark:text-white"
            style={{
              background: isDark ? "rgba(255,255,255,0.04)" : "rgba(0,0,0,0.02)",
            }}
          >
            {isDark ? <Sun className="h-4 w-4" /> : <Moon className="h-4 w-4" />}
          </button>

          {/* User Profile Menu */}
          <div className="relative">
            <button
              onClick={() => setMenuOpen(prev => !prev)}
              className="flex h-9 w-9 items-center justify-center rounded-full transition-all border border-border text-slate-700 dark:text-slate-300 hover:text-foreground dark:text-white hover:border-black/20 dark:border-white/20 active:scale-95"
              style={{
                background: "rgba(255,255,255,0.04)",
                boxShadow: "0 4px 12px -2px rgba(0,0,0,0.3)"
              }}
            >
              <User className="h-4 w-4" />
            </button>

            <AnimatePresence>
              {menuOpen && (
                <>
                  {/* Click outside backdrop to close */}
                  <div className="fixed inset-0 z-40" onClick={() => setMenuOpen(false)} />

                  {/* Dropdown Menu */}
                  <motion.div
                    initial={{ opacity: 0, y: 10, scale: 0.95 }}
                    animate={{ opacity: 1, y: 0, scale: 1 }}
                    exit={{ opacity: 0, y: 10, scale: 0.95 }}
                    transition={{ duration: 0.15 }}
                    className="absolute right-0 mt-2.5 w-64 rounded-2xl border border-border z-50 p-2.5 space-y-1.5 shadow-2xl overflow-hidden"
                    style={{
                      background: "rgba(8, 8, 11, 0.95)",
                      backdropFilter: "blur(20px)"
                    }}
                  >
                    <div className="px-3 py-2 border-b border-border/50 pb-2.5">
                      <p className="text-[10px] text-muted-foreground font-semibold">حساب المستخدم</p>
                      <p className="text-xs font-mono text-foreground dark:text-white truncate mt-0.5">{user.email}</p>
                    </div>

                    <div className="px-3 py-2 bg-black/ dark:bg-black/5 dark:bg-white/[0.02] border border-border/50 rounded-xl">
                      <div className="flex justify-between items-center text-[10px] text-muted-foreground font-semibold">
                        <span>الباقة النشطة:</span>
                        <span className="text-amber-400 uppercase font-extrabold font-mono">Pro Member</span>
                      </div>

                      {plan !== "pro" && (
                        <div className="mt-2.5">
                          <a
                            href="https://tolzy.me/pricing"
                            target="_blank"
                            rel="noopener noreferrer"
                            className="w-full flex items-center justify-center gap-1.5 rounded-lg py-1.5 text-[11px] font-bold text-[#050507] transition-all btn-shimmer mt-2"
                            style={{
                              background: "linear-gradient(135deg, #fff 0%, #e8e8ff 100%)",
                            }}
                            onClick={() => setMenuOpen(false)}
                          >
                            <Sparkles className="h-3 w-3 text-indigo-600" />
                            ترقية إلى Pro
                          </a>
                        </div>
                      )}
                    </div>

                    {/* Sign out option */}
                    <button
                      onClick={() => {
                        setMenuOpen(false);
                        signOut();
                      }}
                      className="w-full flex items-center justify-between text-xs px-3 py-2 rounded-xl text-red-400 hover:text-red-300 hover:bg-red-500/10 transition-colors"
                    >
                      <span>تسجيل الخروج</span>
                      <LogOut className="h-3.5 w-3.5" />
                    </button>
                  </motion.div>
                </>
              )}
            </AnimatePresence>
          </div>

          {/* Upgrade */}
          {plan !== "pro" && (
            <a
              href="https://tolzy.me/pricing"
              target="_blank"
              rel="noopener noreferrer"
              className="relative overflow-hidden text-[11px] font-bold px-4 py-2 rounded-full transition-all btn-shimmer"
              style={{
                background: "linear-gradient(135deg, #6366f1, #818cf8)",
                color: "var(--foreground)",
                boxShadow: "0 4px 20px -4px rgba(99,102,241,0.5)",
              }}
            >
              ترقية إلى Ultra
            </a>
          )}
        </div>
      </nav>
      <section className="relative z-10 mx-auto max-w-5xl px-6 pt-32 pb-8 text-center">

        {/* Status badge */}
        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.5 }}
          className="mx-auto mb-6 inline-flex items-center gap-2.5 rounded-full px-4 py-1.5"
          style={{ background: "rgba(99,102,241,0.1)", border: "1px solid rgba(99,102,241,0.25)" }}
        >
          <div className="glow-dot" />
          <span className="text-[11px] font-semibold" style={{ color: "#a5b4fc" }}>
            Powered by Gemini AI · المعالجة الفورية
          </span>
        </motion.div>

        {/* Main headline */}
        <motion.h1
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6, delay: 0.07 }}
          className="text-4xl sm:text-6xl font-black tracking-tight leading-[1.08]"
        >
          <span className="text-gradient">
            {mode === "excel"
              ? "جداول إكسل ذكية"
              : mode === "presentation"
              ? "عروض تقديمية مبهرة"
              : "مستندات احترافية"}
          </span>
          <br />
          <span className="text-foreground/90 dark:text-white/90">
            {mode === "excel" ? "مدعومة بالذكاء الاصطناعي." : mode === "presentation" ? "مدعومة بالذكاء الاصطناعي." : "مدعومة بالذكاء الاصطناعي."}
          </span>
        </motion.h1>

        {/* Sub headline */}
        <motion.p
          initial={{ opacity: 0, y: 16 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6, delay: 0.14 }}
          className="mt-5 text-base sm:text-lg text-muted-foreground max-w-xl mx-auto leading-relaxed"
        >
          اكتب فكرتك، واحصل على{" "}
          {mode === "excel" ? "جدول Excel" : mode === "presentation" ? "عرض PowerPoint" : "مستند Word"}{" "}
          احترافي جاهز للتحميل في ثوانٍ.
        </motion.p>

        {/* ── Prompt Card ── */}
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6, delay: 0.2 }}
          className="mt-10 relative"
        >
          {/* Mode selector tabs */}
          <div className="flex items-center justify-center gap-1.5 mb-4">
            {(["word", "excel", "presentation"] as DocMode[]).map(m => (
              <button
                key={m}
                onClick={() => { setMode(m); setPrompt(""); }}
                className={`mode-btn ${
                  mode === m
                    ? m === "word" ? "mode-btn-word" : m === "excel" ? "mode-btn-excel" : "mode-btn-ppt"
                    : "mode-btn-inactive"
                }`}
              >
                {m === "excel" ? (
                  <FileSpreadsheet className="h-3.5 w-3.5" />
                ) : m === "presentation" ? (
                  <Presentation className="h-3.5 w-3.5" />
                ) : (
                  <LayoutTemplate className="h-3.5 w-3.5" />
                )}
                {m === "excel" ? "Spreadsheet" : m === "presentation" ? "Presentation" : "Document"}
              </button>
            ))}
          </div>

          {/* Textarea wrapped with elegant soft blue faint ripples */}
          <div className="relative">
            {/* Soft Azure/Indigo Concentric Ambient Glow Rings */}
            <div className="absolute inset-0 flex items-center justify-center pointer-events-none z-0">
              <motion.div
                animate={{ scale: [1, 1.06, 1], opacity: [0.35, 0.55, 0.35] }}
                transition={{ duration: 6, repeat: Infinity, ease: "easeInOut" }}
                className="absolute w-[108%] h-[112%] rounded-3xl blur-[60px]"
                style={{ background: "radial-gradient(circle, rgba(14,165,233,0.14) 0%, rgba(99,102,241,0.04) 50%, transparent 70%)" }}
              />
              <motion.div
                animate={{ scale: [1.05, 0.96, 1.05], opacity: [0.2, 0.4, 0.2] }}
                transition={{ duration: 8, repeat: Infinity, ease: "easeInOut", delay: 1 }}
                className="absolute w-[118%] h-[122%] rounded-3xl blur-[80px]"
                style={{ background: "radial-gradient(circle, rgba(56,189,248,0.08) 0%, transparent 70%)" }}
              />
            </div>

            <div
              className="relative z-10 rounded-3xl overflow-hidden p-6"
              style={{ background: "rgba(255,255,255,0.025)", border: "1px solid rgba(255,255,255,0.09)" }}
            >
              {/* Glow accent line at top */}
              <div className="h-px w-full absolute top-0 left-0" style={{ background: `linear-gradient(90deg, transparent, ${currentMode.color}60, transparent)` }} />

              <textarea
                value={prompt}
                onChange={e => setPrompt(e.target.value.slice(0, 8000))}
                onKeyDown={e => {
                  if (e.key === "Enter" && !e.shiftKey) {
                    e.preventDefault();
                    if (prompt.trim().length >= 8 && !loading) {
                      onGenerate();
                    }
                  }
                }}
                rows={5}
                placeholder={
                  mode === "excel"
                    ? "صف نموذج الجدول الذكي الذي تريد بناءه (مثلاً: نموذج أرباح وخسائر لـ 6 أشهر مع معادلات تلقائية)..."
                    : mode === "presentation"
                    ? "صف موضوع العرض التقديمي الذي تريد بناءه مع الصور والمحتوى..."
                    : "صف المستند أو التقرير الشامل الذي تريد إنشاءه..."
                }
                className="w-full bg-transparent text-xl font-bold text-foreground dark:text-white placeholder:text-muted-foreground outline-none resize-none text-right animate-fade-in"
                style={{ direction: "rtl" }}
              />

              {/* Bottom Row */}
              <div className="flex items-center justify-between mt-4 border-t border-border/50 pt-4">
                {/* Character Count */}
                <div className="text-[10px] text-muted-foreground font-mono">
                  {prompt.length} / 8000
                </div>

                {/* Right side: Model Selector + Mic + Generate Button */}
                <div className="flex items-center gap-3">
                  {/* Model Selector Pill */}
                  <div
                    className="hidden sm:flex items-center gap-1.5 px-3 py-1 rounded-full border border-border/50 text-[11px] font-semibold text-slate-700 dark:text-slate-300 animate-pulse"
                    style={{ background: "rgba(255,255,255,0.02)" }}
                  >
                    <Sparkles className="h-3.5 w-3.5 text-indigo-400" />
                    <span>Gemini 2.0 Flash</span>
                  </div>

                  {/* Submit Action */}
                  <button
                    onClick={onGenerate}
                    disabled={loading || prompt.trim().length < 8}
                    className="flex items-center justify-center gap-2 px-5 py-2.5 rounded-xl font-bold transition-all disabled:opacity-30 disabled:cursor-not-allowed active:scale-95 text-[#050507]"
                    style={{
                      background: loading
                        ? "rgba(255,255,255,0.05)"
                        : "linear-gradient(135deg, #fff 0%, #e8e8ff 100%)",
                    }}
                  >
                    {loading ? (
                      <>
                        <Loader2 className="h-4 w-4 animate-spin text-indigo-600" />
                        <span>جاري البناء...</span>
                      </>
                    ) : (
                      <>
                        <Sparkles className="h-4 w-4 text-indigo-600" />
                        <span>ابدأ التصميم</span>
                      </>
                    )}
                  </button>
                </div>
              </div>
            </div>
          </div>
        </motion.div>


          {/* Quick example chips */}
          <div className="mt-4 flex flex-wrap gap-2 justify-center">
            <span className="text-[10px] text-muted-foreground py-1 font-semibold">💡 جرّب:</span>
            {currentMode.example.map(ex => (
              <button
                key={ex}
                onClick={() => setPrompt(ex)}
                className="tag-pill"
                style={{ direction: "rtl" }}
              >
                {ex}
              </button>
            ))}
          </div>
      </section>

      {/* ── Bento Grid Dashboard ── */}
      <section className="relative z-10 mx-auto max-w-6xl px-6 pb-24 mt-4">
        <div className="grid grid-cols-1 md:grid-cols-12 gap-4 auto-rows-min">

          {/* ── Card 1: Recent Projects (col-span-12) ── */}
          <motion.div
            initial={{ opacity: 0, y: 24 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.3 }}
            onMouseMove={handleMouseMove}
            className="bento-card md:col-span-12 flex flex-col min-h-[280px]"
          >
            {/* Card header */}
            <div className="relative z-10 flex items-center justify-between mb-5">
              <div className="flex items-center gap-3">
                <div className="feature-icon" style={{ background: "rgba(99,102,241,0.15)" }}>
                  <History className="h-5 w-5" style={{ color: "#818cf8" }} />
                </div>
                <div>
                  <h3 className="text-sm font-bold text-foreground dark:text-white">المشاريع الأخيرة</h3>
                  <p className="text-[10px] text-muted-foreground mt-0.5">Recent Flows</p>
                </div>
              </div>
              <span className="text-[10px] font-bold px-2.5 py-1 rounded-full" style={{ background: "rgba(255,255,255,0.06)", color: "#6b6b7b" }}>
                {projects.length} مشروع
              </span>
            </div>

            {/* Project list */}
            <div className="relative z-10 flex-1 space-y-2 overflow-y-auto max-h-[170px] custom-scrollbar pr-1">
              {projects.length === 0 ? (
                <div className="h-full min-h-[120px] flex flex-col items-center justify-center text-center py-8 rounded-2xl" style={{ background: "rgba(255,255,255,0.02)", border: "1px dashed rgba(255,255,255,0.07)" }}>
                  <FolderOpen className="h-8 w-8 mb-2" style={{ color: "#3d3d52" }} />
                  <p className="text-xs font-semibold text-foreground dark:text-white">لا توجد مشاريع بعد</p>
                  <p className="text-[10px] text-muted-foreground mt-1">ابدأ بإنشاء أول مستند الآن!</p>
                </div>
              ) : (
                projects.map(proj => (
                  <div
                    key={proj.id}
                    onClick={() => setActiveProject(proj)}
                    className="group/item flex items-center justify-between p-4 rounded-xl border cursor-pointer transition-all bg-card border-border hover:shadow-md"
                    
                    
                    
                  >
                    <div className="flex items-center gap-3 min-w-0">
                      <div
                        className="flex h-8 w-8 shrink-0 items-center justify-center rounded-xl border transition-all"
                        style={{
                          background: proj.type === "excel"
                            ? "rgba(52,211,153,0.1)"
                            : proj.type === "presentation"
                            ? "rgba(251,146,60,0.1)"
                            : "rgba(129,140,248,0.1)",
                          borderColor: proj.type === "excel"
                            ? "rgba(52,211,153,0.2)"
                            : proj.type === "presentation"
                            ? "rgba(251,146,60,0.2)"
                            : "rgba(129,140,248,0.2)",
                        }}
                      >
                        <ModeIcon mode={proj.type} size="h-3.5 w-3.5" />
                      </div>
                      <div className="min-w-0">
                        <p className="text-xs font-semibold text-foreground dark:text-white truncate group-hover/item:text-indigo-300 transition-colors">{proj.title}</p>
                        <div className="flex items-center gap-1.5 mt-0.5">
                          <Clock className="h-2.5 w-2.5" style={{ color: "#52526a" }} />
                          <p className="text-[9px]" style={{ color: "#52526a" }}>
                            {new Date(proj.updatedAt).toLocaleDateString("ar-EG", { year: "numeric", month: "short", day: "numeric" })}
                          </p>
                        </div>
                      </div>
                    </div>
                    <div className="flex items-center gap-2 shrink-0">
                      <TypeBadge type={proj.type} />
                      <button
                        onClick={(e: any) => handleDeleteProject(proj.id, e)}
                        className="p-1.5 rounded-lg transition-all opacity-0 group-hover/item:opacity-100"
                        style={{ color: "#6b6b7b" }}
                        onMouseEnter={(e: any) => { (e.currentTarget as HTMLElement).style.color = "#ef4444"; (e.currentTarget as HTMLElement).style.background = "rgba(239,68,68,0.1)"; }}
                        onMouseLeave={(e: any) => { (e.currentTarget as HTMLElement).style.color = "#6b6b7b"; (e.currentTarget as HTMLElement).style.background = "transparent"; }}
                      >
                        <Trash2 className="h-3 w-3" />
                      </button>
                    </div>
                  </div>
                ))
              )}
            </div>
          </motion.div>

          {/* ── Feature Cards (3 × col-span-4) ── */}
          {[
            {
              emoji: "📄",
              icon: <FileText className="h-5 w-5" style={{ color: "#818cf8" }} />,
              iconBg: "rgba(129,140,248,0.15)",
              title: "مستندات Word احترافية",
              desc: "هياكل منظمة بعناوين وجداول وفقرات جاهزة للطباعة والتحميل الفوري.",
              stat: "100%",
              statLabel: "جاهز للطباعة",
              color: "#818cf8",
              delay: 0.44,
            },
            {
              emoji: "📊",
              icon: <FileSpreadsheet className="h-5 w-5" style={{ color: "#34d399" }} />,
              iconBg: "rgba(52,211,153,0.15)",
              title: "جداول Excel ذكية",
              desc: "معادلات تلقائية وتنسيقات ألوان احترافية وحسابات دقيقة للميزانيات.",
              stat: "∑",
              statLabel: "معادلات تلقائية",
              color: "#34d399",
              delay: 0.5,
            },
            {
              emoji: "🎥",
              icon: <Presentation className="h-5 w-5" style={{ color: "#fb923c" }} />,
              iconBg: "rgba(251,146,60,0.15)",
              title: "عروض PowerPoint جذابة",
              desc: "شرائح منسقة بجمال فائق مع صور ملائمة ومخصصة من Unsplash.",
              stat: "16:9",
              statLabel: "نسبة احترافية",
              color: "#fb923c",
              delay: 0.56,
            },
          ].map(card => (
            <motion.div
              key={card.title}
              initial={{ opacity: 0, y: 24 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.5, delay: card.delay }}
              onMouseMove={handleMouseMove}
              className="bento-card md:col-span-4 flex flex-col gap-4"
            >
              {/* Icon + stat row */}
              <div className="relative z-10 flex items-start justify-between">
                <div className="feature-icon" style={{ background: card.iconBg }}>
                  {card.icon}
                </div>
                <div className="text-right">
                  <p className="stat-number" style={{ fontSize: "1.5rem" }}>{card.stat}</p>
                  <p className="text-[9px] font-semibold uppercase tracking-widest mt-0.5" style={{ color: card.color }}>{card.statLabel}</p>
                </div>
              </div>

              {/* Text */}
              <div className="relative z-10">
                <h3 className="text-sm font-bold text-foreground dark:text-white mb-1.5">{card.title}</h3>
                <p className="text-[11px] leading-relaxed" style={{ color: "#6b6b7b" }}>{card.desc}</p>
              </div>

              {/* Bottom accent line */}
              <div className="relative z-10 mt-auto h-0.5 w-10 rounded-full" style={{ background: card.color }} />
            </motion.div>
          ))}

        </div>
      </section>

      {/* ── Footer ── */}
      <footer className="relative z-10 border-t py-10 px-6" style={{ borderColor: "rgba(255,255,255,0.05)" }}>
        <div className="max-w-6xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-2.5">
            <div className="flex h-8 w-8 overflow-hidden items-center justify-center rounded-lg border border-border shadow-sm transition-transform duration-300 hover:scale-105 bg-[#0A0A0F]">
              <img src="/image/logo.jpg" alt="TOLZY Flow Logo" className="w-full h-full object-cover" />
            </div>
            <div>
              <p className="text-xs font-black font-mono bg-clip-text text-transparent bg-gradient-to-r from-foreground to-foreground/70 dark:from-white dark:to-white/70">TOLZY Flow</p>
              <p className="text-[9px] text-indigo-400/70 mt-0.5">مهندس المستندات الذكي</p>
            </div>
          </div>
          <p className="text-[10px]" style={{ color: "#52526a" }}>
            © {new Date().getFullYear()} TOLZY Flow · من إنتاج{" "}
            <span className="font-bold text-foreground dark:text-white">TOLZY Labs</span>
          </p>
        </div>
      </footer>
    </main>
  );
}
