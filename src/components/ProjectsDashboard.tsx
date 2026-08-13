import React, { useState, useMemo } from "react";
import { motion, AnimatePresence } from "framer-motion";
import {
  Search,
  Grid,
  List,
  Clock,
  Sparkles,
  Trash2,
  Copy,
  Download,
  ExternalLink,
  FileText,
  FileSpreadsheet,
  Presentation,
  MoreVertical,
  Layers,
  FolderOpen,
  Check,
  Plus,
} from "lucide-react";
import {
  SavedProject,
  deleteProject,
  duplicateProject,
} from "@/lib/projects";
import { generateAndDownloadDocx } from "@/lib/buildDocx";
import { generateAndDownloadExcel } from "@/lib/buildExcel";
import { generateAndDownloadPptx } from "@/lib/buildPptx";
import { toast } from "sonner";

interface ProjectsDashboardProps {
  projects: SavedProject[];
  userId?: string;
  onSelectProject: (id: string) => void;
  onRefreshProjects: () => void;
  onStartNew: (mode?: "word" | "excel" | "presentation") => void;
}

// Normalize type safely so legacy or undefined values never crash
const normalizeType = (type?: string): "word" | "excel" | "presentation" => {
  const t = (type || "").toLowerCase();
  if (t === "excel" || t === "sheet") return "excel";
  if (t === "presentation" || t === "ppt" || t === "pptx") return "presentation";
  return "word";
};

const getModeBadge = (type?: string) => {
  const norm = normalizeType(type);
  switch (norm) {
    case "excel":
      return {
        label: "جدول Excel",
        color: "#34d399",
        bg: "rgba(52,211,153,0.12)",
        border: "rgba(52,211,153,0.3)",
        icon: FileSpreadsheet,
      };
    case "presentation":
      return {
        label: "عرض PPTX",
        color: "#fb923c",
        bg: "rgba(251,146,60,0.12)",
        border: "rgba(251,146,60,0.3)",
        icon: Presentation,
      };
    default:
      return {
        label: "مستند Word",
        color: "#818cf8",
        bg: "rgba(129,140,248,0.12)",
        border: "rgba(129,140,248,0.3)",
        icon: FileText,
      };
  }
};

export function ProjectsDashboard({
  projects = [],
  userId,
  onSelectProject,
  onRefreshProjects,
  onStartNew,
}: ProjectsDashboardProps) {
  const [searchQuery, setSearchQuery] = useState("");
  const [activeFilter, setActiveFilter] = useState<"all" | "word" | "excel" | "presentation">("all");
  const [viewMode, setViewMode] = useState<"grid" | "list">("grid");
  const [sortBy, setSortBy] = useState<"newest" | "oldest" | "title">("newest");
  const [activeMenuId, setActiveMenuId] = useState<string | null>(null);
  const [copiedId, setCopiedId] = useState<string | null>(null);

  // Statistics calculation
  const stats = useMemo(() => {
    const total = projects.length;
    const wordCount = projects.filter((p) => normalizeType(p.type) === "word").length;
    const excelCount = projects.filter((p) => normalizeType(p.type) === "excel").length;
    const pptCount = projects.filter((p) => normalizeType(p.type) === "presentation").length;
    return { total, wordCount, excelCount, pptCount };
  }, [projects]);

  // Filtering & Sorting
  const filteredProjects = useMemo(() => {
    return projects
      .filter((p) => {
        const normType = normalizeType(p.type);
        const matchesFilter = activeFilter === "all" || normType === activeFilter;
        const query = searchQuery.trim().toLowerCase();
        const matchesSearch =
          !query ||
          (p.title && p.title.toLowerCase().includes(query)) ||
          (p.prompt && p.prompt.toLowerCase().includes(query)) ||
          (p.subtitle && p.subtitle.toLowerCase().includes(query));
        return matchesFilter && matchesSearch;
      })
      .sort((a, b) => {
        if (sortBy === "newest") {
          return new Date(b.updatedAt || 0).getTime() - new Date(a.updatedAt || 0).getTime();
        }
        if (sortBy === "oldest") {
          return new Date(a.updatedAt || 0).getTime() - new Date(b.updatedAt || 0).getTime();
        }
        return (a.title || "").localeCompare(b.title || "", "ar");
      });
  }, [projects, activeFilter, searchQuery, sortBy]);

  // Handle Quick Download
  const handleQuickDownload = async (proj: SavedProject, e: React.MouseEvent) => {
    e.stopPropagation();
    try {
      const normType = normalizeType(proj.type);
      if (normType === "excel" && proj.sheets) {
        toast.info("جاري تجهيز وتنزيل ملف Excel...");
        await generateAndDownloadExcel(proj.sheets, proj.title || "مشروع إكسل");
        toast.success("تم تنزيل ملف Excel بنجاح!");
      } else if (normType === "presentation" && (proj.slides || proj.legacySlides)) {
        toast.info("جاري تجهيز وتنزيل العرض التقديمي PPTX...");
        await generateAndDownloadPptx(proj.slides || [], proj.title || "عرض تقديمي", proj.subtitle);
        toast.success("تم تنزيل ملف PPTX بنجاح!");
      } else if (proj.sections) {
        toast.info("جاري تجهيز وتنزيل مستند DOCX...");
        await generateAndDownloadDocx({ title: proj.title || "مستند ورد", subtitle: proj.subtitle, sections: proj.sections });
        toast.success("تم تنزيل مستند DOCX بنجاح!");
      } else {
        toast.error("لا يحتوي هذا المشروع على محتوى قابل للتنزيل المباشر.");
      }
    } catch (err) {
      toast.error("تعذّر التنزيل المباشر، جرب فتح المشروع أولاً.");
    }
  };

  // Handle Duplicate
  const handleDuplicate = async (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    const dup = await duplicateProject(id, userId);
    if (dup) {
      toast.success("تم نسخ المشروع بنجاح!");
      onRefreshProjects();
    }
    setActiveMenuId(null);
  };

  // Handle Delete
  const handleDelete = async (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    await deleteProject(id, userId);
    toast.success("تم حذف المشروع.");
    onRefreshProjects();
    setActiveMenuId(null);
  };

  // Handle Copy Prompt
  const handleCopyPrompt = (proj: SavedProject, e: React.MouseEvent) => {
    e.stopPropagation();
    if (proj.prompt) {
      navigator.clipboard.writeText(proj.prompt);
      setCopiedId(proj.id);
      toast.success("تم نسخ نص الفكرة إلى الحافظة!");
      setTimeout(() => setCopiedId(null), 2000);
    } else {
      toast.info("لا يوجد وصف محفوظ لنسخه.");
    }
    setActiveMenuId(null);
  };

  // Relative Time Helper
  const formatRelativeTime = (dateStr?: string) => {
    if (!dateStr) return "سابقاً";
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return "سابقاً";
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffMins = Math.floor(diffMs / (1000 * 60));
    const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
    const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

    if (diffMins < 2) return "الآن";
    if (diffMins < 60) return `منذ ${diffMins} دقيقة`;
    if (diffHours < 24) return `منذ ${diffHours} ساعة`;
    if (diffDays === 1) return "أمس";
    if (diffDays < 7) return `منذ ${diffDays} أيام`;
    return date.toLocaleDateString("ar-EG", { year: "numeric", month: "short", day: "numeric" });
  };

  return (
    <div className="w-full space-y-6" style={{ direction: "rtl" }}>
      {/* ── TOP STATS & ACTIONS COUNTER BAR ── */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <div className="flex items-center gap-3 p-3.5 rounded-2xl bg-white/[0.03] border border-white/10 backdrop-blur-md">
          <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-indigo-500/15 text-indigo-400 border border-indigo-500/20">
            <Layers className="h-5 w-5" />
          </div>
          <div>
            <p className="text-[11px] text-white/50 font-medium">إجمالي المشاريع</p>
            <p className="text-base font-bold font-mono text-white mt-0.5">{stats.total}</p>
          </div>
        </div>

        <div className="flex items-center gap-3 p-3.5 rounded-2xl bg-white/[0.03] border border-white/10 backdrop-blur-md">
          <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-indigo-500/15 text-indigo-400 border border-indigo-500/20">
            <FileText className="h-5 w-5" />
          </div>
          <div>
            <p className="text-[11px] text-white/50 font-medium">مستندات Word</p>
            <p className="text-base font-bold font-mono text-indigo-300 mt-0.5">{stats.wordCount}</p>
          </div>
        </div>

        <div className="flex items-center gap-3 p-3.5 rounded-2xl bg-white/[0.03] border border-white/10 backdrop-blur-md">
          <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-emerald-500/15 text-emerald-400 border border-emerald-500/20">
            <FileSpreadsheet className="h-5 w-5" />
          </div>
          <div>
            <p className="text-[11px] text-white/50 font-medium">جداول Excel</p>
            <p className="text-base font-bold font-mono text-emerald-300 mt-0.5">{stats.excelCount}</p>
          </div>
        </div>

        <div className="flex items-center gap-3 p-3.5 rounded-2xl bg-white/[0.03] border border-white/10 backdrop-blur-md">
          <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-orange-500/15 text-orange-400 border border-orange-500/20">
            <Presentation className="h-5 w-5" />
          </div>
          <div>
            <p className="text-[11px] text-white/50 font-medium">عروض PPTX</p>
            <p className="text-base font-bold font-mono text-orange-300 mt-0.5">{stats.pptCount}</p>
          </div>
        </div>
      </div>

      {/* ── CONTROLS TOOLBAR (SEARCH + FILTERS + VIEW TOGGLE) ── */}
      <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-4 p-4 rounded-3xl bg-[#0a0a0f]/90 border border-white/10 shadow-xl backdrop-blur-xl">
        {/* Search Bar */}
        <div className="relative flex-1">
          <Search className="absolute right-3.5 top-1/2 -translate-y-1/2 h-4 w-4 text-white/40 pointer-events-none" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="ابحث في مشروعاتك باسم المستند أو الفكرة..."
            className="w-full bg-white/5 text-xs text-white placeholder:text-white/40 pr-10 pl-4 py-2.5 rounded-2xl border border-white/10 outline-none transition-all focus:border-indigo-500/60 focus:bg-white/[0.07]"
          />
          {searchQuery && (
            <button
              onClick={() => setSearchQuery("")}
              className="absolute left-3 top-1/2 -translate-y-1/2 text-xs text-white/50 hover:text-white"
            >
              ✕
            </button>
          )}
        </div>

        {/* Category Filter Chips */}
        <div className="flex items-center gap-1 bg-white/5 p-1 rounded-2xl border border-white/10 overflow-x-auto">
          {[
            { id: "all", label: "الكل", count: stats.total },
            { id: "word", label: "Word 📄", count: stats.wordCount },
            { id: "excel", label: "Excel 📊", count: stats.excelCount },
            { id: "presentation", label: "PPTX 📙", count: stats.pptCount },
          ].map((tab) => (
            <button
              key={tab.id}
              onClick={() => setActiveFilter(tab.id as any)}
              className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all whitespace-nowrap flex items-center gap-1.5 ${
                activeFilter === tab.id
                  ? "bg-white text-black shadow-md scale-[1.02]"
                  : "text-white/70 hover:text-white hover:bg-white/5"
              }`}
            >
              <span>{tab.label}</span>
              <span
                className={`px-1.5 py-0.2 rounded-full text-[10px] font-mono ${
                  activeFilter === tab.id ? "bg-black/10 text-black" : "bg-white/10 text-white/60"
                }`}
              >
                {tab.count}
              </span>
            </button>
          ))}
        </div>

        {/* View mode & Sort */}
        <div className="flex items-center justify-between md:justify-end gap-2 border-t md:border-t-0 border-white/10 pt-3 md:pt-0">
          {/* Sorting */}
          <div className="relative">
            <select
              value={sortBy}
              onChange={(e) => setSortBy(e.target.value as any)}
              className="bg-white/5 text-xs text-white/80 border border-white/10 rounded-xl px-3 py-2 outline-none cursor-pointer hover:bg-white/10 transition-all font-semibold"
            >
              <option value="newest" className="bg-[#0e0e15] text-white">
                الأحدث أولاً
              </option>
              <option value="oldest" className="bg-[#0e0e15] text-white">
                الأقدم أولاً
              </option>
              <option value="title" className="bg-[#0e0e15] text-white">
                أبجدي (الاسم)
              </option>
            </select>
          </div>

          {/* View Toggle */}
          <div className="flex items-center gap-1 bg-white/5 p-1 rounded-xl border border-white/10">
            <button
              onClick={() => setViewMode("grid")}
              className={`p-1.5 rounded-lg transition-all ${
                viewMode === "grid" ? "bg-indigo-500 text-white shadow" : "text-white/50 hover:text-white"
              }`}
              title="عرض شبكي"
            >
              <Grid className="h-4 w-4" />
            </button>
            <button
              onClick={() => setViewMode("list")}
              className={`p-1.5 rounded-lg transition-all ${
                viewMode === "list" ? "bg-indigo-500 text-white shadow" : "text-white/50 hover:text-white"
              }`}
              title="عرض قائمة"
            >
              <List className="h-4 w-4" />
            </button>
          </div>
        </div>
      </div>

      {/* ── PROJECTS CONTENT GRID / LIST ── */}
      {filteredProjects.length === 0 ? (
        <motion.div
          initial={{ opacity: 0, y: 15 }}
          animate={{ opacity: 1, y: 0 }}
          className="flex flex-col items-center justify-center p-12 text-center rounded-3xl border border-dashed border-white/15 bg-white/[0.01]"
        >
          <div className="flex h-16 w-16 items-center justify-center rounded-3xl bg-indigo-500/10 border border-indigo-500/20 text-indigo-400 mb-4">
            <FolderOpen className="h-8 w-8" />
          </div>
          <h3 className="text-base font-bold text-white mb-1">
            {searchQuery ? "لا توجد نتائج تطابق بحثك" : "لا توجد مشاريع محفوظة بعد"}
          </h3>
          <p className="text-xs text-white/50 max-w-sm leading-relaxed mb-6">
            {searchQuery
              ? "جرب البحث بكلمات أخرى أو اختر تصفية مختلفة."
              : "ابدأ بتوليد أول ملف لك الآن بالذكاء الاصطناعي مع الحفظ التلقائي في مكتبتك."}
          </p>
          <button
            onClick={() => onStartNew()}
            className="flex items-center gap-2 px-5 py-2.5 rounded-xl font-bold text-xs bg-white text-black hover:bg-white/90 transition-all shadow-lg hover:scale-105 active:scale-95"
          >
            <Plus className="h-4 w-4 text-indigo-600" />
            إنشاء مستند جديد
          </button>
        </motion.div>
      ) : viewMode === "grid" ? (
        /* GRID VIEW */
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          <AnimatePresence mode="popLayout">
            {filteredProjects.map((proj) => {
              const badge = getModeBadge(proj.type);
              const BadgeIcon = badge.icon;
              const normType = normalizeType(proj.type);

              let summaryInfo = "";
              if (normType === "word" && proj.sections) {
                summaryInfo = `${proj.sections.length} أقسام رئيسية`;
              } else if (normType === "excel" && proj.sheets) {
                summaryInfo = `${proj.sheets.length} أوراق عمل`;
              } else if (normType === "presentation" && proj.slides) {
                summaryInfo = `${proj.slides.length} شرائح العرض`;
              }

              return (
                <motion.div
                  key={proj.id}
                  layout
                  initial={{ opacity: 0, scale: 0.96 }}
                  animate={{ opacity: 1, scale: 1 }}
                  exit={{ opacity: 0, scale: 0.96 }}
                  transition={{ duration: 0.25 }}
                  onClick={() => onSelectProject(proj.id)}
                  className="group relative flex flex-col justify-between rounded-3xl border border-white/10 bg-[#0d0d14]/90 p-5 shadow-lg transition-all duration-300 hover:border-white/25 hover:shadow-2xl hover:-translate-y-1 cursor-pointer overflow-hidden backdrop-blur-xl"
                  style={{
                    boxShadow: `0 15px 35px -10px rgba(0,0,0,0.5)`,
                  }}
                >
                  {/* Subtle Glow Accent Line */}
                  <div
                    className="absolute top-0 inset-x-0 h-1 transition-all duration-300 opacity-60 group-hover:opacity-100"
                    style={{ background: badge.color }}
                  />

                  {/* Top Header Row */}
                  <div>
                    <div className="flex items-center justify-between gap-2 mb-3">
                      <div
                        className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full border text-[11px] font-bold"
                        style={{
                          backgroundColor: badge.bg,
                          borderColor: badge.border,
                          color: badge.color,
                        }}
                      >
                        <BadgeIcon className="h-3.5 w-3.5" />
                        <span>{badge.label}</span>
                      </div>

                      {/* Action Dropdown Menu Button */}
                      <div className="relative">
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            setActiveMenuId(activeMenuId === proj.id ? null : proj.id);
                          }}
                          className="p-1.5 rounded-xl bg-white/5 hover:bg-white/10 text-white/60 hover:text-white transition-all"
                        >
                          <MoreVertical className="h-4 w-4" />
                        </button>

                        {/* Dropdown Menu Popup */}
                        <AnimatePresence>
                          {activeMenuId === proj.id && (
                            <>
                              <div
                                className="fixed inset-0 z-40"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setActiveMenuId(null);
                                }}
                              />
                              <motion.div
                                initial={{ opacity: 0, scale: 0.9, y: 5 }}
                                animate={{ opacity: 1, scale: 1, y: 0 }}
                                exit={{ opacity: 0, scale: 0.9, y: 5 }}
                                transition={{ duration: 0.15 }}
                                className="absolute left-0 mt-1.5 w-44 rounded-2xl border border-white/15 bg-[#12121c]/95 p-1.5 shadow-2xl z-50 space-y-1 text-right backdrop-blur-2xl"
                              >
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    onSelectProject(proj.id);
                                    setActiveMenuId(null);
                                  }}
                                  className="w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs text-white/90 hover:bg-white/10 transition-colors"
                                >
                                  <span>فتح التحرير</span>
                                  <ExternalLink className="h-3.5 w-3.5 text-indigo-400" />
                                </button>

                                <button
                                  onClick={(e) => handleQuickDownload(proj, e)}
                                  className="w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs text-white/90 hover:bg-white/10 transition-colors"
                                >
                                  <span>تنزيل مباشر</span>
                                  <Download className="h-3.5 w-3.5 text-emerald-400" />
                                </button>

                                <button
                                  onClick={(e) => handleCopyPrompt(proj, e)}
                                  className="w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs text-white/90 hover:bg-white/10 transition-colors"
                                >
                                  <span>نسخ الفكرة</span>
                                  {copiedId === proj.id ? (
                                    <Check className="h-3.5 w-3.5 text-emerald-400" />
                                  ) : (
                                    <Copy className="h-3.5 w-3.5 text-amber-400" />
                                  )}
                                </button>

                                <button
                                  onClick={(e) => handleDuplicate(proj.id, e)}
                                  className="w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs text-white/90 hover:bg-white/10 transition-colors"
                                >
                                  <span>تكرار المشروع</span>
                                  <Sparkles className="h-3.5 w-3.5 text-purple-400" />
                                </button>

                                <div className="h-px bg-white/10 my-1" />

                                <button
                                  onClick={(e) => handleDelete(proj.id, e)}
                                  className="w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs text-red-400 hover:bg-red-500/15 transition-colors"
                                >
                                  <span>حذف المشروع</span>
                                  <Trash2 className="h-3.5 w-3.5" />
                                </button>
                              </motion.div>
                            </>
                          )}
                        </AnimatePresence>
                      </div>
                    </div>

                    {/* Title */}
                    <h3 className="text-sm font-bold text-white leading-snug line-clamp-2 group-hover:text-indigo-300 transition-colors">
                      {proj.title || "مشروع بدون عنوان"}
                    </h3>

                    {/* Prompt Preview Snippet */}
                    {proj.prompt && (
                      <p className="text-[11px] text-white/50 line-clamp-2 mt-1.5 leading-relaxed">
                        "{proj.prompt}"
                      </p>
                    )}
                  </div>

                  {/* Card Footer Info */}
                  <div className="flex items-center justify-between mt-5 pt-3 border-t border-white/10 text-[10px] text-white/50">
                    <span className="font-semibold text-white/70">{summaryInfo}</span>
                    <div className="flex items-center gap-1 font-mono">
                      <Clock className="h-3 w-3 text-white/40" />
                      <span>{formatRelativeTime(proj.updatedAt)}</span>
                    </div>
                  </div>
                </motion.div>
              );
            })}
          </AnimatePresence>
        </div>
      ) : (
        /* LIST VIEW */
        <div className="space-y-2">
          <AnimatePresence mode="popLayout">
            {filteredProjects.map((proj) => {
              const badge = getModeBadge(proj.type);
              const BadgeIcon = badge.icon;

              return (
                <motion.div
                  key={proj.id}
                  layout
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: 10 }}
                  onClick={() => onSelectProject(proj.id)}
                  className="group flex items-center justify-between p-4 rounded-2xl border border-white/10 bg-[#0d0d14]/90 hover:border-white/25 hover:bg-white/[0.04] transition-all cursor-pointer shadow-sm"
                >
                  <div className="flex items-center gap-4 min-w-0">
                    <div
                      className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl border"
                      style={{
                        backgroundColor: badge.bg,
                        borderColor: badge.border,
                        color: badge.color,
                      }}
                    >
                      <BadgeIcon className="h-5 w-5" />
                    </div>
                    <div className="min-w-0 text-right">
                      <h4 className="text-xs sm:text-sm font-bold text-white truncate group-hover:text-indigo-300 transition-colors">
                        {proj.title || "مشروع بدون عنوان"}
                      </h4>
                      <p className="text-[11px] text-white/40 truncate max-w-md mt-0.5">
                        {proj.prompt || proj.subtitle || badge.label}
                      </p>
                    </div>
                  </div>

                  <div className="flex items-center gap-3 shrink-0">
                    <span className="hidden sm:inline-block text-[10px] text-white/40 font-mono">
                      {formatRelativeTime(proj.updatedAt)}
                    </span>
                    <button
                      onClick={(e) => handleQuickDownload(proj, e)}
                      className="p-2 rounded-xl bg-white/5 hover:bg-emerald-500/20 text-white/70 hover:text-emerald-300 border border-white/10 transition-all"
                      title="تنزيل مباشر"
                    >
                      <Download className="h-3.5 w-3.5" />
                    </button>
                    <button
                      onClick={(e) => handleDelete(proj.id, e)}
                      className="p-2 rounded-xl bg-white/5 hover:bg-red-500/20 text-white/70 hover:text-red-300 border border-white/10 transition-all"
                      title="حذف"
                    >
                      <Trash2 className="h-3 w-3" />
                    </button>
                  </div>
                </motion.div>
              );
            })}
          </AnimatePresence>
        </div>
      )}
    </div>
  );
}
