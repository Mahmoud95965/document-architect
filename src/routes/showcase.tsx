import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useState } from "react";
import { motion } from "framer-motion";
import { ArrowRight, Sparkles, Sun, Moon, Lock, ExternalLink } from "lucide-react";
import { ShowcaseGallery } from "@/components/ShowcaseGallery";
import { SavedProject } from "@/lib/projects";
import { generateAndDownloadDocx } from "@/lib/buildDocx";
import { generateAndDownloadExcel } from "@/lib/buildExcel";
import { isProjectArabic } from "@/lib/lang";

export const Route = createFileRoute("/showcase")({
  component: ShowcasePage,
  head: () => ({
    meta: [
      { title: "معرض نماذج التوليد الفائقة — TOLZY Flow" },
      {
        name: "description",
        content: "استعرض وعاين نماذج توليد الذكاء الاصطناعي الحية في Word و Excel مجاناً للاطلاع على دقة وسرعة المعالجة.",
      },
    ],
  }),
});

function ShowcasePage() {
  const navigate = useNavigate();
  const [isDark, setIsDark] = useState(true);
  const [activeProject, setActiveProject] = useState<SavedProject | null>(null);
  const [activeSheetIdx, setActiveSheetIdx] = useState(0);

  const toggleTheme = () => {
    setIsDark(!isDark);
    if (!isDark) {
      document.documentElement.classList.add("dark");
    } else {
      document.documentElement.classList.remove("dark");
    }
  };

  // If a project is selected for interactive testing
  if (activeProject) {
    const isExcel = activeProject.type === "excel";
    const isArabic = isProjectArabic(activeProject);
    const sheet = isExcel ? (activeProject.sheets?.[activeSheetIdx] ?? null) : null;

    return (
      <main className="relative min-h-screen bg-background text-foreground flex flex-col" style={{ direction: isArabic ? "rtl" : "ltr" }}>
        {/* Top Header */}
        <header className="h-16 nav-glass fixed top-0 inset-x-0 z-50 flex items-center justify-between px-6 border-b border-border/40">
          <div className="flex items-center gap-3">
            <button
              onClick={() => setActiveProject(null)}
              className="flex items-center gap-1.5 text-xs font-bold text-muted-foreground hover:text-foreground glass px-3 py-1.5 rounded-xl transition-all"
            >
              <ArrowRight className="h-4 w-4 rtl:rotate-180" />
              <span>العودة لمعرض النماذج</span>
            </button>

            <span className="text-xs font-bold text-indigo-600 dark:text-indigo-400 bg-indigo-500/10 px-3 py-1 rounded-xl border border-indigo-500/20">
              معاينة نموذج حية ⚡
            </span>
          </div>

          <div className="flex items-center gap-3">
            <a
              href="https://tolzy.me/pricing"
              target="_blank"
              rel="noopener noreferrer"
              className="flex items-center gap-1.5 text-xs font-bold px-4 py-2 rounded-full text-white bg-indigo-600 hover:bg-indigo-700 shadow-md transition-all btn-shimmer"
            >
              <Sparkles className="h-3.5 w-3.5" />
              <span>ترقية الحساب وإنشاء مستنداتك</span>
            </a>
          </div>
        </header>

        {/* Interactive Viewer Content */}
        <div className="pt-24 pb-16 px-4 sm:px-8 max-w-6xl mx-auto w-full space-y-6">
          <div className="p-6 rounded-3xl bg-white/80 dark:bg-[#0c0c12]/80 border border-border/80 shadow-2xl backdrop-blur-xl space-y-4">
            <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 border-b border-border/60 pb-4">
              <div>
                <span className="text-[10px] font-bold uppercase tracking-wider text-emerald-500 bg-emerald-500/10 px-2.5 py-1 rounded-lg border border-emerald-500/20">
                  {isExcel ? "جدول محاسبي محاكي 📊" : "تقرير Word تنفيذي 📄"}
                </span>
                <h1 className="text-xl sm:text-2xl font-black text-foreground dark:text-white mt-2">
                  {activeProject.title}
                </h1>
                {activeProject.subtitle && (
                  <p className="text-xs text-muted-foreground mt-1">{activeProject.subtitle}</p>
                )}
              </div>

              {/* Direct Download Demo */}
              <button
                onClick={async () => {
                  if (isExcel && activeProject.sheets) {
                    await generateAndDownloadExcel(activeProject.sheets, activeProject.title);
                  } else if (activeProject.sections) {
                    await generateAndDownloadDocx({ title: activeProject.title, subtitle: activeProject.subtitle, sections: activeProject.sections });
                  }
                }}
                className="px-4 py-2 rounded-xl text-xs font-bold text-slate-700 dark:text-slate-200 bg-slate-100 dark:bg-white/10 hover:bg-slate-200 dark:hover:bg-white/20 transition-all border border-border shrink-0"
              >
                تنزيل نسخة تجريبية ({isExcel ? "XLSX" : "DOCX"})
              </button>
            </div>

            {/* Render Preview */}
            {isExcel && activeProject.sheets ? (
              <div className="space-y-4 overflow-x-auto">
                {/* Sheet Tabs */}
                <div className="flex items-center gap-2 border-b border-border/60 pb-2">
                  {activeProject.sheets.map((sh, idx) => (
                    <button
                      key={sh.name}
                      onClick={() => setActiveSheetIdx(idx)}
                      className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all ${
                        activeSheetIdx === idx
                          ? "bg-emerald-500/20 text-emerald-600 dark:text-emerald-300 border border-emerald-500/40 shadow-sm"
                          : "text-muted-foreground hover:bg-slate-100 dark:hover:bg-white/5"
                      }`}
                    >
                      {sh.name}
                    </button>
                  ))}
                </div>

                {/* Grid Table */}
                {sheet && (
                  <div className="border border-border rounded-2xl overflow-hidden shadow-inner bg-white dark:bg-[#08080c]">
                    <table className="w-full text-xs text-right border-collapse">
                      <thead>
                        <tr className="bg-slate-100 dark:bg-white/5 text-slate-700 dark:text-slate-300 font-bold border-b border-border">
                          {sheet.headers.map((h, i) => (
                            <th key={i} className="p-3 border-r border-border font-bold">
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {sheet.rows.map((row, ri) => (
                          <tr key={ri} className="border-b border-border/50 hover:bg-slate-50 dark:hover:bg-white/[0.02]">
                            {sheet.headers.map((_, ci) => {
                              const cell = row[ci] || { value: "" };
                              return (
                                <td
                                  key={ci}
                                  className="p-3 border-r border-border/50 font-mono"
                                  style={{
                                    fontWeight: cell.bold ? "bold" : "normal",
                                    color: cell.color ? `#${cell.color.replace("#", "")}` : undefined,
                                  }}
                                >
                                  {cell.formula ? (
                                    <span className="text-emerald-600 dark:text-emerald-400 font-bold" title={`المعادلة: ${cell.formula}`}>
                                      {cell.value}
                                    </span>
                                  ) : (
                                    String(cell.value ?? "")
                                  )}
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            ) : (
              /* Word Preview */
              <div className="space-y-8 p-4 sm:p-8 bg-white dark:bg-[#09090d] border border-border/60 rounded-2xl">
                {activeProject.sections?.map((sec, idx) => (
                  <div key={idx} className="space-y-3">
                    <h3 className="text-lg font-bold text-foreground dark:text-white border-b border-border/40 pb-2">
                      {sec.heading}
                    </h3>
                    {sec.subheading && (
                      <h4 className="text-xs font-semibold text-indigo-500 dark:text-indigo-400">{sec.subheading}</h4>
                    )}
                    {sec.paragraphs?.map((p, pi) => (
                      <p key={pi} className="text-xs sm:text-sm text-slate-700 dark:text-slate-300 leading-relaxed">
                        {p}
                      </p>
                    ))}
                    {sec.table && (
                      <div className="border border-border rounded-xl overflow-hidden my-4">
                        <table className="w-full text-xs text-right border-collapse">
                          <thead>
                            <tr className="bg-slate-100 dark:bg-white/5 font-bold border-b border-border">
                              {sec.table.headers.map((th, i) => (
                                <th key={i} className="p-2.5 border-r border-border">{th}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {sec.table.rows.map((r, ri) => (
                              <tr key={ri} className="border-b border-border/50">
                                {r.map((c, ci) => (
                                  <td key={ci} className="p-2.5 border-r border-border/50">{c}</td>
                                ))}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      </main>
    );
  }

  return (
    <main className="relative min-h-screen bg-background text-foreground flex flex-col" style={{ direction: "rtl" }}>
      {/* ── NAVBAR ── */}
      <nav className="fixed top-0 inset-x-0 h-16 nav-glass z-50 flex items-center justify-between px-6 border-b border-border/40">
        <div className="flex items-center gap-3">
          <button
            onClick={() => navigate({ to: "/" })}
            className="flex items-center gap-2 text-xs font-bold text-foreground hover:text-indigo-500 transition-colors"
          >
            <div className="flex h-9 w-9 overflow-hidden items-center justify-center rounded-xl border border-indigo-500/20 bg-[#0A0A0F] shadow-sm">
              <img src="/logo.jpg" alt="TOLZY Flow Logo" className="w-full h-full object-cover" />
            </div>
            <span className="text-sm font-black font-mono tracking-wider bg-clip-text text-transparent bg-gradient-to-r from-indigo-600 to-purple-600 dark:from-white dark:to-white/80">
              TOLZY Flow
            </span>
          </button>
        </div>

        <div className="flex items-center gap-3">
          <button
            onClick={toggleTheme}
            className="flex h-9 w-9 items-center justify-center rounded-full transition-all border border-border text-slate-700 dark:text-slate-300 bg-slate-100/80 dark:bg-white/5"
          >
            {isDark ? <Sun className="h-4 w-4 text-amber-400" /> : <Moon className="h-4 w-4 text-indigo-600" />}
          </button>

          <a
            href="https://tolzy.me/pricing"
            target="_blank"
            rel="noopener noreferrer"
            className="flex items-center gap-1.5 text-xs font-bold px-4 py-2 rounded-full text-white bg-indigo-600 hover:bg-indigo-700 shadow-md transition-all btn-shimmer"
          >
            <Sparkles className="h-3.5 w-3.5" />
            <span>ترقية الحساب الآن</span>
          </a>
        </div>
      </nav>

      {/* ── SHOWCASE GALLERY SECTION ── */}
      <div className="pt-24 pb-16">
        <ShowcaseGallery
          onSelectShowcase={(proj) => {
            setActiveProject(proj);
            window.scrollTo({ top: 0, behavior: "smooth" });
          }}
        />
      </div>

      {/* ── FOOTER ── */}
      <footer className="relative z-10 border-t border-border/40 py-8 px-6 bg-slate-50/50 dark:bg-black/30 backdrop-blur-md mt-auto">
        <div className="max-w-6xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-4 text-xs text-muted-foreground">
          <span>© {new Date().getFullYear()} TOLZY Flow · معرض النماذج التفاعلية</span>
          <div>
            <span>من إنتاج وتصميم </span>
            <a
              href="https://tolzy.me"
              target="_blank"
              rel="noopener noreferrer"
              className="font-bold text-indigo-600 dark:text-indigo-400 hover:underline transition-colors"
            >
              TOLZY Team
            </a>
          </div>
        </div>
      </footer>
    </main>
  );
}
