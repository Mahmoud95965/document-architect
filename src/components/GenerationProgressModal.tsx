import React, { useState, useEffect } from "react";
import { motion, AnimatePresence } from "framer-motion";
import {
  Sparkles,
  Brain,
  PenTool,
  Palette,
  CheckCircle2,
  Loader2,
  Lightbulb,
  Shuffle,
  FileText,
  FileSpreadsheet,
  Presentation,
  Zap,
  Clock,
  Wand2,
  Star,
  Layers,
  Layout,
  Table,
} from "lucide-react";

export type DocMode = "word" | "excel" | "presentation";

interface GenerationProgressModalProps {
  isOpen: boolean;
  mode: DocMode;
  prompt: string;
}

// Curated interactive tips & fun facts to keep user engaged while generating
const TASBEER_TIPS = [
  {
    icon: Lightbulb,
    title: "هل تعلم؟",
    text: "التصاميم ذات الهوامش المتناسقة والتباين الواضح تزيد من قراءة واستيعاب المستند بنسبة 45%.",
    tag: "تنسيق المستندات",
  },
  {
    icon: Wand2,
    title: "تلميحة ذكية",
    text: "يمكنك دائماً استخدام المساعد الذكي بعد التوليد للتعديل على أي قسم أو إضافة معادلات جديدة بلغة طبيعية.",
    tag: "مساعد الذكاء الاصطناعي",
  },
  {
    icon: Table,
    title: "معادلات Excel تلقائية",
    text: "يقوم النظام بصياغة صيغ وحسابات المجموع والمتوسطات تلقائياً في جداول البيانات لتوفير وقتك.",
    tag: "جداول البيانات",
  },
  {
    icon: Layout,
    title: "قواعد العروض التقديمية",
    text: "قاعدة (6×6) تنصح بألا تزيد الشريحة عن 6 أسطر، وكل سطر 6 كلمات للحفاظ على تركيز الجمهور.",
    tag: "العروض التقديمية",
  },
  {
    icon: Zap,
    title: "الإنتاجية الفائقة",
    text: "توليد هيكل الملف بالذكاء الاصطناعي يوفر في المتوسط 85% من الوقت المستغرق في التنسيق اليدوي.",
    tag: "سرعة الإنجاز",
  },
  {
    icon: Star,
    title: "تصدير مباشر",
    text: "يمكنك تصدير النتائج مباشرة بصيغ DOCX أو XLSX أو PPTX مع الحفاظ الكامل على التنسيق والأنماط.",
    tag: "التصدير الاحترافي",
  },
  {
    icon: Brain,
    title: "تحليل ذكي للمحتوى",
    text: "يعتمد محرك Axiom على قراءة تفاصيل وصفك بعناية لبناء هيكل متكامل يناسب طبيعة عملك.",
    tag: "ذكاء الاصطناعي",
  },
];

interface Particle {
  id: number;
  x: number;
  y: number;
  size: number;
  color: string;
}

export function GenerationProgressModal({ isOpen, mode, prompt }: GenerationProgressModalProps) {
  const [currentTipIndex, setCurrentTipIndex] = useState(0);
  const [currentStep, setCurrentStep] = useState(0);
  const [progress, setProgress] = useState(5);
  const [elapsedSeconds, setElapsedSeconds] = useState(0);
  const [particles, setParticles] = useState<Particle[]>([]);

  // Mode Theme Configuration
  const modeConfig = {
    word: {
      name: "مستند Word احترافي",
      color: "#6366f1", // Indigo
      gradient: "from-indigo-500 to-purple-600",
      bgGlow: "rgba(99, 102, 241, 0.18)",
      borderColor: "rgba(99, 102, 241, 0.3)",
      badgeBg: "rgba(99, 102, 241, 0.12)",
      icon: FileText,
      blueprintIcon: Layers,
    },
    excel: {
      name: "جدول بيانات Excel ذكي",
      color: "#10b981", // Emerald
      gradient: "from-emerald-500 to-teal-600",
      bgGlow: "rgba(16, 185, 129, 0.18)",
      borderColor: "rgba(16, 185, 129, 0.3)",
      badgeBg: "rgba(16, 185, 129, 0.12)",
      icon: FileSpreadsheet,
      blueprintIcon: Table,
    },
    presentation: {
      name: "عرض تقديمي Presentation جذّاب",
      color: "#f97316", // Orange
      gradient: "from-orange-500 to-amber-600",
      bgGlow: "rgba(249, 115, 22, 0.18)",
      borderColor: "rgba(249, 115, 22, 0.3)",
      badgeBg: "rgba(249, 115, 22, 0.12)",
      icon: Presentation,
      blueprintIcon: Layout,
    },
  };

  const currentTheme = modeConfig[mode] || modeConfig.word;

  const STEPS = [
    { title: "تحليل المتطلبات وهيكلة الأفكار", desc: "قراءة الوصف وبناء المخطط الهيكلي الأساسي", icon: Brain },
    { title: "صياغة المحتوى وتوليد البيانات", desc: "إنشاء النصوص والتفاصيل والمعادلات الدقيقة", icon: PenTool },
    { title: "تنسيق التصميم والألوان والخطوط", desc: "تطبيق التنسيقات البصرية والجماليات الحديثة", icon: Palette },
    { title: "تجهيز المستند النهائي والتأكد من الجودة", desc: "اللمسات الأخيرة وتحضير ملف التصدير", icon: CheckCircle2 },
  ];

  // Timer & Progress simulation
  useEffect(() => {
    if (!isOpen) {
      setCurrentStep(0);
      setProgress(5);
      setElapsedSeconds(0);
      return;
    }

    // Elapsed timer
    const timerInterval = setInterval(() => {
      setElapsedSeconds((prev) => prev + 1);
    }, 1000);

    // Progress bar and step transitions
    const progressInterval = setInterval(() => {
      setProgress((prev) => {
        if (prev >= 94) return 94; // Wait for full server finish
        const increment = Math.max(0.5, (95 - prev) * 0.08);
        return prev + increment;
      });
    }, 200);

    // Step updater based on progress
    const stepInterval = setInterval(() => {
      setProgress((p) => {
        if (p < 25) setCurrentStep(0);
        else if (p < 55) setCurrentStep(1);
        else if (p < 82) setCurrentStep(2);
        else setCurrentStep(3);
        return p;
      });
    }, 300);

    // Auto rotate tips
    const tipInterval = setInterval(() => {
      setCurrentTipIndex((prev) => (prev + 1) % TASBEER_TIPS.length);
    }, 4000);

    return () => {
      clearInterval(timerInterval);
      clearInterval(progressInterval);
      clearInterval(stepInterval);
      clearInterval(tipInterval);
    };
  }, [isOpen]);

  // Click interaction to spawn particles (Tasbeer fun game)
  const handleSparkleClick = (e: React.MouseEvent<HTMLDivElement>) => {
    const rect = e.currentTarget.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    const newParticle: Particle = {
      id: Date.now() + Math.random(),
      x,
      y,
      size: Math.floor(Math.random() * 12) + 8,
      color: currentTheme.color,
    };

    setParticles((prev) => [...prev.slice(-15), newParticle]);
  };

  const nextTip = () => {
    setCurrentTipIndex((prev) => (prev + 1) % TASBEER_TIPS.length);
  };

  const prevTip = () => {
    setCurrentTipIndex((prev) => (prev - 1 + TASBEER_TIPS.length) % TASBEER_TIPS.length);
  };

  if (!isOpen) return null;

  const currentTip = TASBEER_TIPS[currentTipIndex];
  const TipIcon = currentTip.icon;
  const ModeIcon = currentTheme.icon;

  return (
    <AnimatePresence>
      <div
        className="fixed inset-0 z-[120] flex items-center justify-center p-4 sm:p-6 overflow-y-auto bg-black/85 backdrop-blur-xl selection:bg-indigo-500/30"
        style={{ direction: "rtl" }}
      >
        {/* Background Ambient Glow */}
        <div className="absolute inset-0 overflow-hidden pointer-events-none">
          <div
            className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[700px] h-[700px] rounded-full blur-[140px] opacity-35 transition-all duration-700"
            style={{ background: currentTheme.bgGlow }}
          />
          <div className="absolute inset-0 bg-[radial-gradient(circle_at_center,rgba(255,255,255,0.03)_1px,transparent_1px)] bg-[size:24px_24px] opacity-40" />
        </div>

        {/* Modal Container */}
        <motion.div
          initial={{ opacity: 0, scale: 0.94, y: 20 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.94, y: 20 }}
          transition={{ type: "spring", stiffness: 300, damping: 28 }}
          className="relative w-full max-w-2xl rounded-3xl border border-white/10 bg-[#0c0c12]/95 p-6 sm:p-8 shadow-2xl overflow-hidden backdrop-blur-2xl"
          style={{
            boxShadow: `0 25px 70px -15px rgba(0,0,0,0.8), 0 0 50px ${currentTheme.bgGlow}`,
          }}
        >
          {/* Header Bar */}
          <div className="flex items-center justify-between pb-5 border-b border-white/10">
            <div className="flex items-center gap-3">
              <div
                className="flex h-11 w-11 items-center justify-center rounded-2xl border"
                style={{
                  backgroundColor: currentTheme.badgeBg,
                  borderColor: currentTheme.borderColor,
                }}
              >
                <ModeIcon className="h-5.5 w-5.5" style={{ color: currentTheme.color }} />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-bold tracking-wide uppercase text-white/60">
                    جاري التوليد بواسطة الذكاء الاصطناعي
                  </span>
                  <span
                    className="inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-[10px] font-bold"
                    style={{
                      backgroundColor: currentTheme.badgeBg,
                      color: currentTheme.color,
                    }}
                  >
                    <Sparkles className="h-3 w-3 animate-spin" />
                    Axiom Engine
                  </span>
                </div>
                <h3 className="text-lg font-bold text-white mt-0.5">{currentTheme.name}</h3>
              </div>
            </div>

            {/* Elapsed Time Counter */}
            <div className="flex items-center gap-1.5 rounded-full bg-white/5 px-3 py-1.5 border border-white/10 text-xs font-mono text-white/80">
              <Clock className="h-3.5 w-3.5 text-indigo-400 animate-pulse" />
              <span>{elapsedSeconds} ثوانٍ</span>
            </div>
          </div>

          {/* Interactive Sparkle Playground Blueprint Visualizer */}
          <div
            onClick={handleSparkleClick}
            className="relative my-6 rounded-2xl border border-white/10 bg-black/40 p-5 overflow-hidden cursor-pointer group select-none transition-all hover:border-white/20"
          >
            {/* Background scanner line effect */}
            <motion.div
              className="absolute inset-x-0 h-1 opacity-40 pointer-events-none"
              style={{
                background: `linear-gradient(90deg, transparent 0%, ${currentTheme.color} 50%, transparent 100%)`,
              }}
              animate={{ y: [0, 160, 0] }}
              transition={{ duration: 3.5, repeat: Infinity, ease: "easeInOut" }}
            />

            {/* Floating spawned particles on click */}
            {particles.map((p) => (
              <motion.div
                key={p.id}
                initial={{ opacity: 1, scale: 0.2, x: p.x, y: p.y }}
                animate={{ opacity: 0, scale: 1.8, y: p.y - 45 }}
                transition={{ duration: 0.8, ease: "easeOut" }}
                className="absolute pointer-events-none flex items-center justify-center"
                style={{ left: 0, top: 0 }}
              >
                <Sparkles className="h-5 w-5" style={{ color: p.color }} />
              </motion.div>
            ))}

            <div className="relative z-10 flex flex-col sm:flex-row items-center justify-between gap-4">
              {/* Dynamic Animated Blueprint Card */}
              <div className="flex items-center gap-4">
                <div className="relative flex h-14 w-14 shrink-0 items-center justify-center rounded-2xl bg-white/5 border border-white/10">
                  <motion.div
                    className="absolute inset-0 rounded-2xl border-2 border-transparent"
                    style={{
                      borderTopColor: currentTheme.color,
                      borderRightColor: "transparent",
                    }}
                    animate={{ rotate: 360 }}
                    transition={{ duration: 2, repeat: Infinity, ease: "linear" }}
                  />
                  <currentTheme.blueprintIcon
                    className="h-7 w-7 transition-all duration-300"
                    style={{ color: currentTheme.color }}
                  />
                </div>
                <div className="text-right">
                  <p className="text-xs text-white/50 font-medium">الوصف المطلوب:</p>
                  <p className="text-sm font-semibold text-white/90 line-clamp-1 max-w-sm">
                    "{prompt}"
                  </p>
                  <p className="text-[11px] text-white/40 mt-1 flex items-center gap-1">
                    <span>✨ اضغط هنا في أي مكان لإطلاق شرارات النجوم!</span>
                  </p>
                </div>
              </div>

              {/* Dynamic Percentage Pill */}
              <div className="flex flex-col items-center sm:items-end shrink-0">
                <div className="text-2xl font-black font-mono text-white tracking-tight">
                  {Math.round(progress)}%
                </div>
                <span className="text-[11px] text-white/50 font-medium">معدل البناء</span>
              </div>
            </div>

            {/* Smooth Progress Bar */}
            <div className="mt-4 h-2.5 w-full rounded-full bg-white/10 overflow-hidden p-0.5 border border-white/5">
              <motion.div
                className="h-full rounded-full transition-all duration-300"
                style={{
                  width: `${progress}%`,
                  background: `linear-gradient(90deg, ${currentTheme.color}, #a855f7)`,
                  boxShadow: `0 0 12px ${currentTheme.color}`,
                }}
              />
            </div>
          </div>

          {/* Stepper Progress Section */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-6">
            {STEPS.map((step, idx) => {
              const isDone = idx < currentStep;
              const isCurrent = idx === currentStep;
              const StepIcon = step.icon;

              return (
                <div
                  key={step.title}
                  className={`relative flex items-start gap-3 p-3 rounded-2xl border transition-all duration-300 ${
                    isCurrent
                      ? "bg-white/[0.07] border-white/20 shadow-lg"
                      : isDone
                      ? "bg-white/[0.02] border-emerald-500/30 opacity-90"
                      : "bg-white/[0.01] border-white/5 opacity-40"
                  }`}
                >
                  <div
                    className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-xl transition-all ${
                      isDone
                        ? "bg-emerald-500/20 text-emerald-400 border border-emerald-500/40"
                        : isCurrent
                        ? "bg-indigo-500/20 text-indigo-400 border border-indigo-500/40 animate-pulse"
                        : "bg-white/5 text-white/40 border border-white/10"
                    }`}
                  >
                    {isDone ? (
                      <CheckCircle2 className="h-4.5 w-4.5 text-emerald-400" />
                    ) : isCurrent ? (
                      <Loader2 className="h-4.5 w-4.5 animate-spin text-indigo-400" />
                    ) : (
                      <StepIcon className="h-4 w-4" />
                    )}
                  </div>
                  <div className="text-right">
                    <h4
                      className={`text-xs font-bold transition-colors ${
                        isCurrent
                          ? "text-white"
                          : isDone
                          ? "text-emerald-300"
                          : "text-white/50"
                      }`}
                    >
                      {step.title}
                    </h4>
                    <p className="text-[11px] text-white/50 mt-0.5 leading-tight">{step.desc}</p>
                  </div>
                </div>
              );
            })}
          </div>

          {/* ── TASBEER KNOWLEDGE TICKER CAROUSEL (نظام التصبير والتسلية المفيدة) ── */}
          <div className="relative rounded-2xl border border-white/10 bg-gradient-to-r from-indigo-950/30 via-purple-950/20 to-slate-950/40 p-4 sm:p-5">
            <div className="flex items-center justify-between mb-2">
              <div className="flex items-center gap-2">
                <div className="flex h-6 w-6 items-center justify-center rounded-lg bg-amber-500/20 text-amber-300 border border-amber-500/30">
                  <Lightbulb className="h-3.5 w-3.5" />
                </div>
                <span className="text-xs font-bold text-amber-300">💡 فائدة أثناء الانتظار</span>
                <span className="rounded-full bg-white/10 px-2 py-0.5 text-[10px] font-semibold text-white/60">
                  {currentTip.tag}
                </span>
              </div>

              {/* Controls */}
              <div className="flex items-center gap-1.5">
                <button
                  onClick={prevTip}
                  className="p-1 rounded-lg bg-white/5 hover:bg-white/10 text-white/70 hover:text-white transition-all"
                  title="السابق"
                >
                  ‹
                </button>
                <button
                  onClick={nextTip}
                  className="flex items-center gap-1 px-2.5 py-1 rounded-lg bg-indigo-500/20 hover:bg-indigo-500/30 text-indigo-300 border border-indigo-500/30 text-[11px] font-semibold transition-all"
                >
                  <Shuffle className="h-3 w-3" />
                  <span>تلميحة جديدة</span>
                </button>
                <button
                  onClick={nextTip}
                  className="p-1 rounded-lg bg-white/5 hover:bg-white/10 text-white/70 hover:text-white transition-all"
                  title="التالي"
                >
                  ›
                </button>
              </div>
            </div>

            {/* Tip Body */}
            <div className="min-h-[44px] flex items-center">
              <AnimatePresence mode="wait">
                <motion.div
                  key={currentTipIndex}
                  initial={{ opacity: 0, y: 8 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -8 }}
                  transition={{ duration: 0.3 }}
                  className="flex items-start gap-3 w-full"
                >
                  <TipIcon className="h-5 w-5 shrink-0 text-amber-400 mt-0.5" />
                  <p className="text-xs sm:text-sm leading-relaxed text-white/90 font-medium text-right">
                    {currentTip.text}
                  </p>
                </motion.div>
              </AnimatePresence>
            </div>
          </div>
        </motion.div>
      </div>
    </AnimatePresence>
  );
}
