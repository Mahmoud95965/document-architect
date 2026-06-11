import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useState, useEffect } from "react";
import { 
  signInWithEmailAndPassword, 
  GoogleAuthProvider, 
  signInWithPopup, 
  createUserWithEmailAndPassword 
} from "firebase/auth";
import { doc, setDoc } from "firebase/firestore";
import { motion, AnimatePresence } from "framer-motion";
import { Sparkles, Mail, Lock, Loader2, AlertCircle, Eye, EyeOff, ArrowRight, LogOut } from "lucide-react";
import { auth, db } from "@/lib/firebase";
import { useAuth } from "@/hooks/useAuth";


export const Route = createFileRoute("/login")({
  component: LoginPage,
  head: () => ({
    meta: [
      { title: "تسجيل الدخول — AXIOM Flow" },
      { name: "description", content: "سجّل دخولك إلى منصة AXIOM لتوليد وتصميم المستندات والشرائح الذكية بالذكاء الاصطناعي." },
      { property: "og:title", content: "تسجيل الدخول — AXIOM Flow" },
      { property: "og:description", content: "سجّل دخولك للوصول إلى منصة AXIOM لتوليد ملفات Word وExcel وPowerPoint بالذكاء الاصطناعي." },
    ],
  }),
});

const GoogleIcon = () => (
  <svg className="h-4 w-4" viewBox="0 0 24 24" style={{ display: "inline-block" }}>
    <path
      fill="#EA4335"
      d="M12 5.04c1.66 0 3.2.57 4.38 1.69l3.27-3.27C17.67 1.48 14.97 1 12 1 7.35 1 3.4 3.65 1.5 7.5l3.82 2.96C6.23 7.38 8.87 5.04 12 5.04z"
    />
    <path
      fill="#4285F4"
      d="M23.49 12.27c0-.81-.07-1.59-.2-2.36H12v4.51h6.43c-.28 1.44-1.1 2.66-2.33 3.47l3.62 2.81c2.12-1.95 3.77-4.83 3.77-8.43z"
    />
    <path
      fill="#FBBC05"
      d="M5.32 10.46a7.16 7.16 0 0 1 0 3.08l-3.82 2.96A11.96 11.96 0 0 1 1.5 7.5l3.82 2.96z"
    />
    <path
      fill="#34A853"
      d="M12 23c3.24 0 5.97-1.07 7.96-2.91l-3.62-2.81c-1.1.74-2.5 1.18-4.34 1.18-3.13 0-5.77-2.34-6.68-5.42L1.5 16.5C3.4 20.35 7.35 23 12 23z"
    />
  </svg>
);

function LoginPage() {
  const { user, plan, loading: authLoading, signOut } = useAuth();
  const navigate = useNavigate();

  const [activeTab,       setActiveTab]       = useState<"login" | "register">("login");
  const [email,           setEmail]           = useState("");
  const [password,        setPassword]        = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [showPass,        setShowPass]        = useState(false);
  const [showConfirmPass, setShowConfirmPass] = useState(false);
  const [loading,         setLoading]         = useState(false);
  const [error,           setError]           = useState<string | null>(null);


  // Already logged in as pro → go straight to app
  useEffect(() => {
    if (!authLoading && user && plan === "pro") {
      navigate({ to: "/" });
    }
  }, [authLoading, user, plan, navigate]);

  async function handleLogin(e: React.FormEvent) {
    e.preventDefault();
    if (!email.trim() || !password) return;

    setLoading(true);
    setError(null);

    try {
      await signInWithEmailAndPassword(auth, email.trim(), password);
      // onAuthStateChanged in useAuth will update plan → redirect handled by useEffect above
    } catch (err: unknown) {
      const code = (err as { code?: string }).code ?? "";
      if (
        code === "auth/user-not-found" ||
        code === "auth/wrong-password" ||
        code === "auth/invalid-credential" ||
        code === "auth/invalid-email"
      ) {
        setError("البريد الإلكتروني أو كلمة المرور غير صحيحة.");
      } else if (code === "auth/too-many-requests") {
        setError("تم إيقاف الحساب مؤقتاً بسبب محاولات كثيرة. حاول لاحقاً.");
      } else if (code === "auth/network-request-failed") {
        setError("تعذّر الاتصال بالشبكة. تحقق من الإنترنت وأعد المحاولة.");
      } else {
        setError("حدث خطأ غير متوقع. حاول مجدداً.");
      }
    } finally {
      setLoading(false);
    }
  }  async function handleRegister(e: React.FormEvent) {
    e.preventDefault();
    if (!email.trim() || !password || !confirmPassword) return;

    if (password.length < 6) {
      setError("يجب أن تكون كلمة المرور 6 أحرف على الأقل.");
      return;
    }

    if (password !== confirmPassword) {
      setError("كلمتا المرور غير متطابقتين.");
      return;
    }

    setLoading(true);
    setError(null);

    try {
      const userCredential = await createUserWithEmailAndPassword(auth, email.trim(), password);
      const registeredUser = userCredential.user;

      // Create a user document in Firestore with 'free' plan
      await setDoc(doc(db, "users", registeredUser.uid), {
        email: registeredUser.email,
        plan: "free",
        createdAt: new Date().toISOString(),
        filesCountToday: 0,
        lastGenerationDate: new Date().toISOString().split("T")[0],
      }, { merge: true });
    } catch (err: unknown) {
      const code = (err as { code?: string }).code ?? "";
      if (code === "auth/email-already-in-use") {
        setError("البريد الإلكتروني هذا مستخدم بالفعل.");
      } else if (code === "auth/invalid-email") {
        setError("البريد الإلكتروني غير صالح.");
      } else if (code === "auth/weak-password") {
        setError("كلمة المرور ضعيفة جداً. يجب أن تكون 6 أحرف على الأقل.");
      } else if (code === "auth/network-request-failed") {
        setError("تعذّر الاتصال بالشبكة. تحقق من الإنترنت وأعد المحاولة.");
      } else {
        setError("حدث خطأ أثناء إنشاء الحساب. حاول مجدداً.");
      }
    } finally {
      setLoading(false);
    }
  }

  function handleSubmit(e: React.FormEvent) {
    if (activeTab === "login") {
      handleLogin(e);
    } else {
      handleRegister(e);
    }
  }

  async function handleGoogleLogin() {
    setLoading(true);
    setError(null);
    try {
      const provider = new GoogleAuthProvider();
      await signInWithPopup(auth, provider);
    } catch (err: unknown) {
      const code = (err as { code?: string }).code ?? "";
      if (code === "auth/popup-closed-by-user") {
        setError("تم إغلاق نافذة تسجيل الدخول.");
      } else if (code === "auth/network-request-failed") {
        setError("تعذّر الاتصال بالشبكة. تحقق من الإنترنت وأعد المحاولة.");
      } else {
        setError("حدث خطأ أثناء تسجيل الدخول بواسطة Google.");
      }
    } finally {
      setLoading(false);
    }
  }

  const isNonPro = user && plan !== "pro";

  // Show a full-screen spinner while auth state is resolving
  if (authLoading) {
    return (
      <div className="min-h-screen bg-[#050507] flex items-center justify-center">
        <Loader2 className="h-8 w-8 animate-spin text-indigo-400" />
      </div>
    );
  }

  return (
    <main className="relative min-h-screen bg-[#050507] flex flex-col items-center justify-center px-4 overflow-hidden">

      {/* Dot grid */}
      <div className="pointer-events-none absolute inset-0 grid-bg opacity-40" />

      {/* Glow blobs */}
      <div className="pointer-events-none absolute inset-0 overflow-hidden">
        <div
          className="hero-glow-blob"
          style={{
            width: "600px", height: "600px",
            top: "-200px", left: "-150px",
            background: "radial-gradient(circle, rgba(99,102,241,0.16) 0%, transparent 70%)",
          }}
        />
        <div
          className="hero-glow-blob"
          style={{
            width: "400px", height: "400px",
            bottom: "0px", right: "-100px",
            background: "radial-gradient(circle, rgba(249,115,22,0.10) 0%, transparent 70%)",
            animationDelay: "-4s",
          }}
        />
      </div>

      {isNonPro ? (
        /* ── Upgrade Card ── */
        <motion.div
          initial={{ opacity: 0, y: 28, scale: 0.97 }}
          animate={{ opacity: 1, y: 0, scale: 1 }}
          transition={{ duration: 0.55, ease: [0.16, 1, 0.3, 1] }}
          className="relative z-10 w-full max-w-md"
        >
          {/* Top accent bar - Premium Gold */}
          <div className="h-px w-full mb-0" style={{ background: "linear-gradient(90deg, transparent, rgba(245,158,11,0.6), transparent)" }} />

          <div
            className="rounded-3xl p-8 sm:p-10 text-center"
            style={{
              background: "linear-gradient(145deg, rgba(255,255,255,0.04) 0%, rgba(255,255,255,0.01) 100%)",
              border: "1px solid rgba(255,255,255,0.08)",
              boxShadow: "0 40px 100px -30px rgba(0,0,0,0.9), 0 0 0 1px rgba(255,255,255,0.03) inset",
            }}
          >
            {/* Logo + Icon */}
            <div className="flex flex-col items-center text-center mb-6">
              <div
                className="relative flex h-16 w-16 items-center justify-center rounded-2xl mb-5"
                style={{ background: "linear-gradient(135deg, #2a1f10, #13131a)", border: "1px solid rgba(245, 158, 11, 0.25)" }}
              >
                <Sparkles className="h-7 w-7 text-amber-400 animate-pulse" />
                {/* Animated ring */}
                <div
                  className="absolute inset-0 rounded-2xl border-glow"
                  style={{ opacity: 0.5, borderColor: "rgba(245, 158, 11, 0.3)" }}
                />
              </div>

              {/* Badge */}
              <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-amber-500/10 text-amber-300 border border-amber-500/20 mb-4" style={{ direction: "rtl" }}>
                ✨ حساب عادي / غير نشط
              </span>

              <h1 className="text-2xl font-black text-white tracking-tight">مطلوب ترقية الحساب إلى Pro</h1>
              <p className="mt-3 text-sm leading-relaxed" style={{ color: "#a1a1b5", direction: "rtl" }}>
                عذراً، حسابك الحالي لا يمتلك اشتراك <span className="text-amber-300 font-semibold">TOLZY Pro</span> نشط. لتتمكن من إنشاء وتوليد مستندات Word وجداول Excel وعروض PowerPoint التفاعلية، يرجى ترقية خطتك.
              </p>
            </div>

            {/* Email info */}
            <div className="rounded-xl py-2 px-4 mb-6 bg-white/5 border border-white/5 inline-block text-xs" style={{ color: "#8e8e9f", direction: "rtl" }}>
              حسابك الحالي: <span className="text-white font-mono font-semibold">{user?.email}</span>
            </div>

            {/* Actions */}
            <div className="space-y-3">
              <a
                href="https://tolzy.me/pricing"
                target="_blank"
                rel="noopener noreferrer"
                className="w-full flex items-center justify-center gap-2.5 rounded-2xl py-3.5 text-sm font-bold text-[#050507] transition-all btn-shimmer"
                style={{
                  background: "linear-gradient(135deg, #f59e0b 0%, #d97706 100%)",
                  boxShadow: "0 4px 20px -4px rgba(245,158,11,0.4)",
                }}
              >
                <Sparkles className="h-4 w-4 text-[#050507]" />
                ترقية الحساب الآن
              </a>

              <button
                onClick={() => signOut()}
                className="w-full flex items-center justify-center gap-2.5 rounded-2xl py-3.5 text-sm font-semibold transition-all"
                style={{
                  background: "rgba(255,255,255,0.04)",
                  border: "1px solid rgba(255,255,255,0.08)",
                  color: "#fff",
                }}
                onMouseEnter={e => { e.currentTarget.style.background = "rgba(255,255,255,0.08)"; }}
                onMouseLeave={e => { e.currentTarget.style.background = "rgba(255,255,255,0.04)"; }}
              >
                <LogOut className="h-4 w-4 text-gray-400" />
                <span>تسجيل الخروج أو تبديل الحساب</span>
              </button>
            </div>
          </div>

          {/* Powered-by note */}
          <p className="mt-5 text-center text-[10px]" style={{ color: "#3d3d52" }}>
            مشغّل بواسطة <span className="text-[#52526a] font-semibold">TOLZY Labs</span> · جميع الحقوق محفوظة
          </p>
        </motion.div>
      ) : (
        /* ── Login Card ── */
        <motion.div
          initial={{ opacity: 0, y: 28, scale: 0.97 }}
          animate={{ opacity: 1, y: 0, scale: 1 }}
          transition={{ duration: 0.55, ease: [0.16, 1, 0.3, 1] }}
          className="relative z-10 w-full max-w-md animate-float-up"
        >
          {/* Top accent bar - Dynamic color depending on activeTab */}
          <div 
            className="h-px w-full mb-0 transition-all duration-500" 
            style={{ 
              background: activeTab === "login" 
                ? "linear-gradient(90deg, transparent, rgba(99,102,241,0.6), transparent)" 
                : "linear-gradient(90deg, transparent, rgba(16,185,129,0.6), transparent)" 
            }} 
          />

          <div
            className="rounded-3xl p-8 sm:p-10 relative overflow-hidden"
            style={{
              background: "linear-gradient(145deg, rgba(255,255,255,0.04) 0%, rgba(255,255,255,0.01) 100%)",
              border: "1px solid rgba(255,255,255,0.08)",
              boxShadow: "0 40px 100px -30px rgba(0,0,0,0.9), 0 0 0 1px rgba(255,255,255,0.03) inset",
            }}
          >
            {/* Ambient inner glow */}
            <div 
              className="absolute top-0 right-0 w-[150px] h-[150px] rounded-full blur-[80px] pointer-events-none transition-all duration-500" 
              style={{ 
                background: activeTab === "login"
                  ? "radial-gradient(circle, rgba(99,102,241,0.08) 0%, transparent 70%)"
                  : "radial-gradient(circle, rgba(16,185,129,0.08) 0%, transparent 70%)"
              }} 
            />

            {/* Logo + heading */}
            <div className="flex flex-col items-center text-center mb-6 relative z-10">
              <div
                className="relative flex h-16 w-16 items-center justify-center rounded-2xl mb-5 transition-all duration-500"
                style={{ 
                  background: "linear-gradient(135deg, #1e1e2e 0%, #0c0c14 100%)", 
                  border: activeTab === "login" 
                    ? "1px solid rgba(129, 140, 248, 0.25)" 
                    : "1px solid rgba(52, 211, 153, 0.25)",
                  boxShadow: activeTab === "login"
                    ? "0 10px 30px -5px rgba(99, 102, 241, 0.2)"
                    : "0 10px 30px -5px rgba(16, 185, 129, 0.2)"
                }}
              >
                <Sparkles className={`h-7 w-7 transition-colors duration-500 animate-pulse ${activeTab === "login" ? "text-indigo-400" : "text-emerald-400"}`} />
                {/* Animated ring */}
                <div
                  className="absolute inset-0 rounded-2xl border-glow"
                  style={{ opacity: 0.6 }}
                />
              </div>

              <h1 className="text-2xl font-black text-white tracking-tight bg-clip-text text-transparent bg-gradient-to-r from-white via-white to-slate-400 mb-2">
                {activeTab === "login" ? "تسجيل الدخول" : "إنشاء حساب جديد"}
              </h1>
              <p className="text-xs" style={{ color: "#8f8f9f" }}>
                {activeTab === "login" ? "سجّل دخولك للوصول إلى منصة" : "أنشئ حساباً مجانياً للبدء في استخدام"}{" "}
                <span className="font-extrabold" style={{ color: activeTab === "login" ? "#a5b4fc" : "#6ee7b7" }}>AXIOM Flow</span>
              </p>
            </div>

            {/* Tabs switcher */}
            <div 
              className="flex p-1 rounded-2xl mb-6 relative z-10"
              style={{ 
                background: "rgba(255, 255, 255, 0.02)",
                border: "1px solid rgba(255, 255, 255, 0.05)",
                width: "100%",
                direction: "rtl"
              }}
            >
              <button
                type="button"
                onClick={() => { setActiveTab("login"); setError(null); }}
                className="flex-1 py-2.5 text-xs font-bold rounded-xl transition-all relative cursor-pointer"
                style={{
                  color: activeTab === "login" ? "#fff" : "#71717a",
                }}
              >
                {activeTab === "login" && (
                  <motion.div
                    layoutId="active-tab-bg"
                    className="absolute inset-0 rounded-xl"
                    style={{
                      background: "linear-gradient(135deg, rgba(99,102,241,0.12) 0%, rgba(79,70,229,0.18) 100%)",
                      border: "1px solid rgba(129, 140, 248, 0.2)",
                    }}
                    transition={{ type: "spring", stiffness: 380, damping: 30 }}
                  />
                )}
                <span className="relative z-10">تسجيل الدخول</span>
              </button>

              <button
                type="button"
                onClick={() => { setActiveTab("register"); setError(null); }}
                className="flex-1 py-2.5 text-xs font-bold rounded-xl transition-all relative cursor-pointer"
                style={{
                  color: activeTab === "register" ? "#fff" : "#71717a",
                }}
              >
                {activeTab === "register" && (
                  <motion.div
                    layoutId="active-tab-bg"
                    className="absolute inset-0 rounded-xl"
                    style={{
                      background: "linear-gradient(135deg, rgba(16,185,129,0.12) 0%, rgba(5,150,105,0.18) 100%)",
                      border: "1px solid rgba(52, 211, 153, 0.2)",
                    }}
                    transition={{ type: "spring", stiffness: 380, damping: 30 }}
                  />
                )}
                <span className="relative z-10">إنشاء حساب</span>
              </button>
            </div>

            {/* Form */}
            <form onSubmit={handleSubmit} className="space-y-5 relative z-10" noValidate>

              {/* Email */}
              <div className="space-y-2">
                <label className="block text-xs font-semibold text-right" style={{ color: "#9898a8", direction: "rtl" }}>
                  البريد الإلكتروني
                </label>
                <div className="relative">
                  <Mail
                    className="absolute left-4 top-1/2 -translate-y-1/2 h-4 w-4 pointer-events-none transition-colors duration-300"
                    style={{ color: "#52526a" }}
                  />
                  <input
                    id="tolzy-email"
                    type="email"
                    autoComplete="email"
                    value={email}
                    onChange={e => { setEmail(e.target.value); setError(null); }}
                    placeholder="name@example.com"
                    required
                    className="w-full rounded-2xl pl-11 pr-4 py-3.5 text-sm text-white placeholder:text-[#3d3d52] outline-none transition-all duration-300"
                    style={{
                      background: "rgba(255,255,255,0.02)",
                      border: "1px solid rgba(255,255,255,0.06)",
                    }}
                    onFocus={e => { 
                      e.target.style.borderColor = activeTab === "login" ? "rgba(129,140,248,0.4)" : "rgba(52,211,153,0.4)"; 
                      e.target.style.background = "rgba(255,255,255,0.04)";
                      e.target.style.boxShadow = activeTab === "login" ? "0 0 20px -3px rgba(99,102,241,0.15)" : "0 0 20px -3px rgba(16,185,129,0.15)";
                    }}
                    onBlur={e => { 
                      e.target.style.borderColor = "rgba(255,255,255,0.06)"; 
                      e.target.style.background = "rgba(255,255,255,0.02)";
                      e.target.style.boxShadow = "none";
                    }}
                  />
                </div>
              </div>

              {/* Password */}
              <div className="space-y-2">
                <label className="block text-xs font-semibold text-right" style={{ color: "#9898a8", direction: "rtl" }}>
                  كلمة المرور
                </label>
                <div className="relative">
                  <Lock
                    className="absolute left-4 top-1/2 -translate-y-1/2 h-4 w-4 pointer-events-none transition-colors duration-300"
                    style={{ color: "#52526a" }}
                  />
                  <input
                    id="tolzy-password"
                    type={showPass ? "text" : "password"}
                    autoComplete={activeTab === "login" ? "current-password" : "new-password"}
                    value={password}
                    onChange={e => { setPassword(e.target.value); setError(null); }}
                    placeholder="••••••••••"
                    required
                    className="w-full rounded-2xl pl-11 pr-11 py-3.5 text-sm text-white placeholder:text-[#3d3d52] outline-none transition-all duration-300"
                    style={{
                      background: "rgba(255,255,255,0.02)",
                      border: "1px solid rgba(255,255,255,0.06)",
                    }}
                    onFocus={e => { 
                      e.target.style.borderColor = activeTab === "login" ? "rgba(129,140,248,0.4)" : "rgba(52,211,153,0.4)"; 
                      e.target.style.background = "rgba(255,255,255,0.04)";
                      e.target.style.boxShadow = activeTab === "login" ? "0 0 20px -3px rgba(99,102,241,0.15)" : "0 0 20px -3px rgba(16,185,129,0.15)";
                    }}
                    onBlur={e => { 
                      e.target.style.borderColor = "rgba(255,255,255,0.06)"; 
                      e.target.style.background = "rgba(255,255,255,0.02)";
                      e.target.style.boxShadow = "none";
                    }}
                  />
                  <button
                    type="button"
                    onClick={() => setShowPass(v => !v)}
                    className="absolute right-4 top-1/2 -translate-y-1/2 transition-colors cursor-pointer"
                    style={{ color: "#52526a" }}
                    onMouseEnter={e => (e.currentTarget.style.color = activeTab === "login" ? "#818cf8" : "#34d399")}
                    onMouseLeave={e => (e.currentTarget.style.color = "#52526a")}
                  >
                    {showPass ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
                  </button>
                </div>
              </div>

              {/* Confirm Password (only for register) */}
              <AnimatePresence>
                {activeTab === "register" && (
                  <motion.div 
                    key="confirm-password"
                    initial={{ opacity: 0, height: 0, marginTop: 0 }}
                    animate={{ opacity: 1, height: "auto", marginTop: 8 }}
                    exit={{ opacity: 0, height: 0, marginTop: 0 }}
                    transition={{ duration: 0.25, ease: "easeInOut" }}
                    className="space-y-2 overflow-hidden"
                  >
                    <label className="block text-xs font-semibold text-right" style={{ color: "#9898a8", direction: "rtl" }}>
                      تأكيد كلمة المرور
                    </label>
                    <div className="relative">
                      <Lock
                        className="absolute left-4 top-1/2 -translate-y-1/2 h-4 w-4 pointer-events-none transition-colors duration-300"
                        style={{ color: "#52526a" }}
                      />
                      <input
                        id="tolzy-confirm-password"
                        type={showConfirmPass ? "text" : "password"}
                        autoComplete="new-password"
                        value={confirmPassword}
                        onChange={e => { setConfirmPassword(e.target.value); setError(null); }}
                        placeholder="••••••••••"
                        required
                        className="w-full rounded-2xl pl-11 pr-11 py-3.5 text-sm text-white placeholder:text-[#3d3d52] outline-none transition-all duration-300"
                        style={{
                          background: "rgba(255,255,255,0.02)",
                          border: "1px solid rgba(255,255,255,0.06)",
                        }}
                        onFocus={e => { 
                          e.target.style.borderColor = "rgba(52,211,153,0.4)"; 
                          e.target.style.background = "rgba(255,255,255,0.04)";
                          e.target.style.boxShadow = "0 0 20px -3px rgba(16,185,129,0.15)";
                        }}
                        onBlur={e => { 
                          e.target.style.borderColor = "rgba(255,255,255,0.06)"; 
                          e.target.style.background = "rgba(255,255,255,0.02)";
                          e.target.style.boxShadow = "none";
                        }}
                      />
                      <button
                        type="button"
                        onClick={() => setShowConfirmPass(v => !v)}
                        className="absolute right-4 top-1/2 -translate-y-1/2 transition-colors cursor-pointer"
                        style={{ color: "#52526a" }}
                        onMouseEnter={e => (e.currentTarget.style.color = "#34d399")}
                        onMouseLeave={e => (e.currentTarget.style.color = "#52526a")}
                      >
                        {showConfirmPass ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
                      </button>
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>

              {/* Error message */}
              {error && (
                <motion.div
                  initial={{ opacity: 0, y: -6 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="flex items-start gap-2.5 rounded-2xl px-4 py-3.5 text-right"
                  style={{ background: "rgba(239,68,68,0.06)", border: "1px solid rgba(239,68,68,0.15)", direction: "rtl" }}
                >
                  <AlertCircle className="h-4 w-4 text-red-400 shrink-0 mt-0.5" />
                  <p className="text-xs text-red-300 leading-relaxed">{error}</p>
                </motion.div>
              )}

              {/* Submit */}
              <button
                id="login-submit-btn"
                type="submit"
                disabled={loading || !email.trim() || !password || (activeTab === "register" && !confirmPassword)}
                className="w-full flex items-center justify-center gap-2.5 rounded-2xl py-3.5 text-sm font-bold text-white transition-all disabled:opacity-40 disabled:cursor-not-allowed btn-shimmer mt-3 shadow-lg cursor-pointer"
                style={{
                  background: activeTab === "login" 
                    ? "linear-gradient(135deg, #6366f1 0%, #4f46e5 100%)" 
                    : "linear-gradient(135deg, #10b981 0%, #059669 100%)",
                  boxShadow: loading 
                    ? "none" 
                    : activeTab === "login" 
                      ? "0 4px 25px -5px rgba(99,102,241,0.4)"
                      : "0 4px 25px -5px rgba(16,185,129,0.4)",
                }}
              >
                {loading ? (
                  <Loader2 className="h-4 w-4 animate-spin" />
                ) : activeTab === "login" ? (
                  <>
                    <Sparkles className="h-4 w-4 text-indigo-200" />
                    تسجيل الدخول للمنصة
                    <ArrowRight className="h-4 w-4 text-indigo-200" />
                  </>
                ) : (
                  <>
                    <Sparkles className="h-4 w-4 text-emerald-200" />
                    إنشاء حساب جديد
                    <ArrowRight className="h-4 w-4 text-emerald-200" />
                  </>
                )}
              </button>
            </form>

            {/* Divider */}
            <div className="mt-6 flex items-center gap-3 relative z-10">
              <div className="flex-1 h-px" style={{ background: "rgba(255,255,255,0.05)" }} />
              <span className="text-[10px] font-semibold uppercase tracking-widest text-[#3d3d52]">أو</span>
              <div className="flex-1 h-px" style={{ background: "rgba(255,255,255,0.05)" }} />
            </div>

            {/* Google Sign In */}
            <button
              onClick={handleGoogleLogin}
              disabled={loading}
              className="w-full flex items-center justify-center gap-2.5 rounded-2xl py-3.5 text-sm font-semibold transition-all disabled:opacity-40 mt-4 cursor-pointer relative z-10"
              style={{
                background: "rgba(255,255,255,0.02)",
                border: "1px solid rgba(255,255,255,0.06)",
                color: "#fff",
              }}
              onMouseEnter={e => { 
                e.currentTarget.style.background = "rgba(255,255,255,0.05)"; 
                e.currentTarget.style.borderColor = "rgba(255,255,255,0.1)";
              }}
              onMouseLeave={e => { 
                e.currentTarget.style.background = "rgba(255,255,255,0.02)"; 
                e.currentTarget.style.borderColor = "rgba(255,255,255,0.06)";
              }}
            >
              {loading ? (
                <Loader2 className="h-4 w-4 animate-spin" />
              ) : (
                <>
                  <GoogleIcon />
                  <span>متابعة بواسطة Google</span>
                </>
              )}
            </button>

            {/* Footer note switcher */}
            <p className="mt-6 text-center text-[12px] leading-relaxed relative z-10" style={{ color: "#52526a", direction: "rtl" }}>
              {activeTab === "login" ? (
                <>
                  لا تملك حساباً؟{" "}
                  <button
                    type="button"
                    onClick={() => { setActiveTab("register"); setError(null); }}
                    className="font-bold transition-colors cursor-pointer"
                    style={{ color: "#818cf8" }}
                    onMouseEnter={e => (e.currentTarget.style.color = "#a5b4fc")}
                    onMouseLeave={e => (e.currentTarget.style.color = "#818cf8")}
                  >
                    سجل حساباً جديداً من هنا ←
                  </button>
                </>
              ) : (
                <>
                  لديك حساب بالفعل؟{" "}
                  <button
                    type="button"
                    onClick={() => { setActiveTab("login"); setError(null); }}
                    className="font-bold transition-colors cursor-pointer"
                    style={{ color: "#818cf8" }}
                    onMouseEnter={e => (e.currentTarget.style.color = "#a5b4fc")}
                    onMouseLeave={e => (e.currentTarget.style.color = "#818cf8")}
                  >
                    سجّل دخولك من هنا ←
                  </button>
                </>
              )}
            </p>
          </div>

          {/* Powered-by note */}
          <p className="mt-5 text-center text-[10px] relative z-10" style={{ color: "#3d3d52" }}>
            مشغّل بواسطة <span className="text-[#52526a] font-semibold">AXIOM Labs</span> · جميع الحقوق محفوظة
          </p>
        </motion.div>

      )}
    </main>
  );
}
