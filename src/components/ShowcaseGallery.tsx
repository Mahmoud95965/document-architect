import React from "react";
import { motion } from "framer-motion";
import { Sparkles, FileSpreadsheet, FileText, Eye, CheckCircle2, ArrowLeft } from "lucide-react";
import { SHOWCASE_PROJECTS } from "@/lib/showcase.data";
import { SavedProject } from "@/lib/projects";

interface ShowcaseGalleryProps {
  onSelectShowcase: (project: SavedProject) => void;
}

export function ShowcaseGallery({ onSelectShowcase }: ShowcaseGalleryProps) {
  return (
    <section className="relative z-10 mx-auto max-w-6xl px-6 py-12" style={{ direction: "rtl" }}>
      {/* ── SECTION HEADER ── */}
      <div className="text-center max-w-3xl mx-auto mb-10 space-y-3">
        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          className="inline-flex items-center gap-2 px-4 py-1.5 rounded-full border border-indigo-500/30 bg-indigo-500/10 text-indigo-700 dark:text-indigo-300 text-xs font-bold shadow-sm backdrop-blur-md"
        >
          <Sparkles className="h-3.5 w-3.5 text-indigo-500 animate-pulse" />
          <span>معرض أعمال وتوليد الذكاء الاصطناعي (Showcase)</span>
        </motion.div>

        <h2 className="text-2xl sm:text-3xl font-black text-foreground dark:text-white tracking-tight">
          شاهد قوة الدقة والتنسيق الفائق قبل الترقية ⚡
        </h2>

        <p className="text-xs sm:text-sm text-muted-foreground leading-relaxed">
          اختر أي من النماذج الحية التالية للاطلاع على دقة المعادلات، ثراء الفقرات، وتناسق الجداول المنشأة بمحرك الذكاء الاصطناعي مباشرة.
        </p>
      </div>

      {/* ── SHOWCASE BENTO GRID ── */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        {SHOWCASE_PROJECTS.map((project, idx) => {
          const isExcel = project.type === "excel";
          const color = isExcel ? "#34d399" : "#818cf8";
          const Icon = isExcel ? FileSpreadsheet : FileText;

          return (
            <motion.div
              key={project.id}
              initial={{ opacity: 0, y: 24 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.5, delay: idx * 0.15 }}
              className="group relative rounded-3xl border border-border/80 bg-white/80 dark:bg-[#0c0c12]/80 p-6 shadow-xl backdrop-blur-xl transition-all duration-300 hover:scale-[1.01] hover:border-indigo-500/40 flex flex-col justify-between"
            >
              {/* Background ambient glow */}
              <div
                className="absolute inset-0 rounded-3xl opacity-0 group-hover:opacity-100 transition-opacity duration-500 pointer-events-none"
                style={{
                  background: `radial-gradient(500px circle at 50% 0%, ${color}15, transparent 70%)`,
                }}
              />

              <div className="relative z-10 space-y-4">
                {/* Top badge row */}
                <div className="flex items-center justify-between gap-2">
                  <div
                    className="flex items-center gap-2 px-3 py-1 rounded-xl text-xs font-bold border shadow-sm"
                    style={{
                      backgroundColor: `${color}15`,
                      color: color,
                      borderColor: `${color}30`,
                    }}
                  >
                    <Icon className="h-4 w-4" />
                    <span>{isExcel ? "نماذج جداول Excel حية" : "تقرير Word تنفيذي دسم"}</span>
                  </div>

                  <span className="text-[10px] font-mono font-bold text-muted-foreground flex items-center gap-1 bg-slate-100 dark:bg-white/5 px-2.5 py-1 rounded-lg border border-border">
                    <CheckCircle2 className="h-3 w-3 text-emerald-500" />
                    معادلات وهياكل حية 100%
                  </span>
                </div>

                {/* Title & Description */}
                <div>
                  <h3 className="text-lg font-bold text-foreground dark:text-white group-hover:text-indigo-600 dark:group-hover:text-indigo-400 transition-colors">
                    {project.title}
                  </h3>
                  <p className="text-xs text-muted-foreground mt-1.5 leading-relaxed">
                    {project.subtitle}
                  </p>
                </div>

                {/* Live Data Features Chips */}
                <div className="flex flex-wrap gap-1.5 py-1">
                  {isExcel ? (
                    <>
                      <span className="text-[10px] font-semibold bg-emerald-500/10 text-emerald-600 dark:text-emerald-300 border border-emerald-500/20 px-2.5 py-1 rounded-lg">
                        📊 2 ورقة عمل متكاملة (Sheets)
                      </span>
                      <span className="text-[10px] font-semibold bg-indigo-500/10 text-indigo-600 dark:text-indigo-300 border border-indigo-500/20 px-2.5 py-1 rounded-lg">
                        ∑ معادلات =SUM و =IF و =TODAY
                      </span>
                      <span className="text-[10px] font-semibold bg-orange-500/10 text-orange-600 dark:text-orange-300 border border-orange-500/20 px-2.5 py-1 rounded-lg">
                        🎨 ألوان حالات وحسابات العجز
                      </span>
                    </>
                  ) : (
                    <>
                      <span className="text-[10px] font-semibold bg-indigo-500/10 text-indigo-600 dark:text-indigo-300 border border-indigo-500/20 px-2.5 py-1 rounded-lg">
                        📄 7 أقسام تحليلية كاملة
                      </span>
                      <span className="text-[10px] font-semibold bg-emerald-500/10 text-emerald-600 dark:text-emerald-300 border border-emerald-500/20 px-2.5 py-1 rounded-lg">
                        📋 جداول مقارنة وتأثير القطاعات
                      </span>
                      <span className="text-[10px] font-semibold bg-purple-500/10 text-purple-600 dark:text-purple-300 border border-purple-500/20 px-2.5 py-1 rounded-lg">
                        ✍️ صياغة لغوية ودراسة اقتصادية
                      </span>
                    </>
                  )}
                </div>

                {/* Sample Content Snippet Box */}
                <div className="p-3.5 rounded-2xl bg-slate-50 dark:bg-black/40 border border-border text-xs text-muted-foreground space-y-1.5 font-mono overflow-hidden max-h-28 relative">
                  <div className="absolute inset-x-0 bottom-0 h-10 bg-gradient-to-t from-slate-50 dark:from-[#08080c] to-transparent pointer-events-none" />
                  {isExcel ? (
                    <div>
                      <p className="font-bold text-foreground dark:text-slate-200">الورقة الأولى: {project.sheets?.[0]?.name}</p>
                      <p className="text-[11px] text-emerald-600 dark:text-emerald-400 font-bold mt-1">
                        Headers: {project.sheets?.[0]?.headers.join(" | ")}
                      </p>
                      <p className="text-[10px] text-slate-500 mt-1">
                        أول عنصر: {String(project.sheets?.[0]?.rows[0][1]?.value)} — معادلة الحالة: {project.sheets?.[0]?.rows[0][7]?.formula}
                      </p>
                    </div>
                  ) : (
                    <div>
                      <p className="font-bold text-foreground dark:text-slate-200">{project.sections?.[0]?.heading}</p>
                      <p className="text-[11px] line-clamp-2 mt-1">
                        {project.sections?.[0]?.paragraphs?.[0]}
                      </p>
                    </div>
                  )}
                </div>
              </div>

              {/* Action Button */}
              <div className="relative z-10 pt-5 mt-4 border-t border-border/60">
                <button
                  onClick={() => onSelectShowcase(project)}
                  className="w-full flex items-center justify-center gap-2 py-3 px-4 rounded-2xl text-xs font-bold text-white transition-all duration-300 shadow-md hover:shadow-indigo-500/25 active:scale-[0.98] btn-shimmer"
                  style={{
                    background: isExcel
                      ? "linear-gradient(135deg, #059669 0%, #10b981 100%)"
                      : "linear-gradient(135deg, #4f46e5 0%, #6366f1 100%)",
                  }}
                >
                  <Eye className="h-4 w-4" />
                  <span>معاينة واختبار النموذج التفاعلي</span>
                  <ArrowLeft className="h-3.5 w-3.5 rtl:rotate-180" />
                </button>
              </div>
            </motion.div>
          );
        })}
      </div>
    </section>
  );
}
