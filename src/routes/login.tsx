import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useState, useEffect } from "react";
import { signInWithEmailAndPassword, GoogleAuthProvider, signInWithPopup } from "firebase/auth";
import { motion } from "framer-motion";
import { Sparkles, Mail, Lock, Loader2, AlertCircle, Eye, EyeOff, ArrowRight } from "lucide-react";
import { auth } from "@/lib/firebase";
import { useAuth } from "@/hooks/useAuth";

export const Route = createFileRoute("/login")({
  component: LoginPage,
  head: () => ({
    meta: [
      { title: "تسجيل الدخول — TOLZY Flow" },
      { name: "description", content: "سجّل دخولك للوصول إلى TOLZY Flow." },
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
  const { user, plan, loading: authLoading } = useAuth();
  const navigate = useNavigate();

  const [email,       setEmail]       = useState("");
  const [password,    setPassword]    = useState("");
  const [showPass,    setShowPass]    = useState(false);
  const [loading,     setLoading]     = useState(false);
  const [error,       setError]       = useState<string | null>(null);

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

      {/* ── Login Card ── */}
      <motion.div
        initial={{ opacity: 0, y: 28, scale: 0.97 }}
        animate={{ opacity: 1, y: 0, scale: 1 }}
        transition={{ duration: 0.55, ease: [0.16, 1, 0.3, 1] }}
        className="relative z-10 w-full max-w-md"
      >
        {/* Top accent bar */}
        <div className="h-px w-full mb-0" style={{ background: "linear-gradient(90deg, transparent, rgba(99,102,241,0.6), transparent)" }} />

        <div
          className="rounded-3xl p-8 sm:p-10"
          style={{
            background: "linear-gradient(145deg, rgba(255,255,255,0.04) 0%, rgba(255,255,255,0.01) 100%)",
            border: "1px solid rgba(255,255,255,0.08)",
            boxShadow: "0 40px 100px -30px rgba(0,0,0,0.9), 0 0 0 1px rgba(255,255,255,0.03) inset",
          }}
        >
          {/* Logo + heading */}
          <div className="flex flex-col items-center text-center mb-8">
            <div
              className="relative flex h-16 w-16 items-center justify-center rounded-2xl mb-5"
              style={{ background: "linear-gradient(135deg, #1e1e2e, #13131a)", border: "1px solid rgba(255,255,255,0.1)" }}
            >
              <Sparkles className="h-7 w-7 text-indigo-400" />
              {/* Animated ring */}
              <div
                className="absolute inset-0 rounded-2xl border-glow"
                style={{ opacity: 0.5 }}
              />
            </div>

            <h1 className="text-2xl font-black text-white tracking-tight">تسجيل الدخول</h1>
            <p className="mt-1.5 text-sm" style={{ color: "#6b6b7b" }}>
              سجّل دخولك للوصول إلى{" "}
              <span className="font-semibold" style={{ color: "#a5b4fc" }}>TOLZY Flow</span>
            </p>
          </div>

          {/* Form */}
          <form onSubmit={handleLogin} className="space-y-4" noValidate>

            {/* Email */}
            <div className="space-y-1.5">
              <label className="block text-xs font-semibold text-right" style={{ color: "#9898a8", direction: "rtl" }}>
                البريد الإلكتروني
              </label>
              <div className="relative">
                <Mail
                  className="absolute left-4 top-1/2 -translate-y-1/2 h-4 w-4 pointer-events-none"
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
                  className="w-full rounded-xl pl-11 pr-4 py-3 text-sm text-white placeholder:text-[#3d3d52] outline-none transition-all"
                  style={{
                    background: "rgba(255,255,255,0.04)",
                    border: "1px solid rgba(255,255,255,0.08)",
                  }}
                  onFocus={e => { e.target.style.borderColor = "rgba(129,140,248,0.5)"; e.target.style.background = "rgba(255,255,255,0.06)"; }}
                  onBlur={e => { e.target.style.borderColor = "rgba(255,255,255,0.08)"; e.target.style.background = "rgba(255,255,255,0.04)"; }}
                />
              </div>
            </div>

            {/* Password */}
            <div className="space-y-1.5">
              <label className="block text-xs font-semibold text-right" style={{ color: "#9898a8", direction: "rtl" }}>
                كلمة المرور
              </label>
              <div className="relative">
                <Lock
                  className="absolute left-4 top-1/2 -translate-y-1/2 h-4 w-4 pointer-events-none"
                  style={{ color: "#52526a" }}
                />
                <input
                  id="tolzy-password"
                  type={showPass ? "text" : "password"}
                  autoComplete="current-password"
                  value={password}
                  onChange={e => { setPassword(e.target.value); setError(null); }}
                  placeholder="••••••••••"
                  required
                  className="w-full rounded-xl pl-11 pr-11 py-3 text-sm text-white placeholder:text-[#3d3d52] outline-none transition-all"
                  style={{
                    background: "rgba(255,255,255,0.04)",
                    border: "1px solid rgba(255,255,255,0.08)",
                  }}
                  onFocus={e => { e.target.style.borderColor = "rgba(129,140,248,0.5)"; e.target.style.background = "rgba(255,255,255,0.06)"; }}
                  onBlur={e => { e.target.style.borderColor = "rgba(255,255,255,0.08)"; e.target.style.background = "rgba(255,255,255,0.04)"; }}
                />
                <button
                  type="button"
                  onClick={() => setShowPass(v => !v)}
                  className="absolute right-4 top-1/2 -translate-y-1/2 transition-colors"
                  style={{ color: "#52526a" }}
                  onMouseEnter={e => (e.currentTarget.style.color = "#9898a8")}
                  onMouseLeave={e => (e.currentTarget.style.color = "#52526a")}
                >
                  {showPass ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
                </button>
              </div>
            </div>

            {/* Error message */}
            {error && (
              <motion.div
                initial={{ opacity: 0, y: -6 }}
                animate={{ opacity: 1, y: 0 }}
                className="flex items-start gap-2.5 rounded-xl px-4 py-3 text-right"
                style={{ background: "rgba(239,68,68,0.08)", border: "1px solid rgba(239,68,68,0.2)", direction: "rtl" }}
              >
                <AlertCircle className="h-4 w-4 text-red-400 shrink-0 mt-0.5" />
                <p className="text-xs text-red-300 leading-relaxed">{error}</p>
              </motion.div>
            )}

            {/* Submit */}
            <button
              id="login-submit-btn"
              type="submit"
              disabled={loading || !email.trim() || !password}
              className="w-full flex items-center justify-center gap-2.5 rounded-2xl py-3 text-sm font-bold text-[#050507] transition-all disabled:opacity-40 disabled:cursor-not-allowed btn-shimmer mt-2"
              style={{
                background: "linear-gradient(135deg, #fff 0%, #e8e8ff 100%)",
                boxShadow: loading ? "none" : "0 4px 24px -6px rgba(255,255,255,0.25)",
              }}
            >
              {loading ? (
                <Loader2 className="h-4 w-4 animate-spin" />
              ) : (
                <>
                  <Sparkles className="h-4 w-4 text-indigo-600" />
                  تسجيل الدخول
                  <ArrowRight className="h-4 w-4 text-indigo-600" />
                </>
              )}
            </button>
          </form>

          {/* Divider */}
          <div className="mt-6 flex items-center gap-3">
            <div className="flex-1 h-px" style={{ background: "rgba(255,255,255,0.05)" }} />
            <span className="text-[10px] font-semibold uppercase tracking-widest" style={{ color: "#3d3d52" }}>أو</span>
            <div className="flex-1 h-px" style={{ background: "rgba(255,255,255,0.05)" }} />
          </div>

          {/* Google Sign In */}
          <button
            onClick={handleGoogleLogin}
            disabled={loading}
            className="w-full flex items-center justify-center gap-2.5 rounded-2xl py-3 text-sm font-semibold transition-all disabled:opacity-40 mt-4"
            style={{
              background: "rgba(255,255,255,0.04)",
              border: "1px solid rgba(255,255,255,0.08)",
              color: "#fff",
            }}
            onMouseEnter={e => { e.currentTarget.style.background = "rgba(255,255,255,0.08)"; }}
            onMouseLeave={e => { e.currentTarget.style.background = "rgba(255,255,255,0.04)"; }}
          >
            {loading ? (
              <Loader2 className="h-4 w-4 animate-spin" />
            ) : (
              <>
                <GoogleIcon />
                <span>تسجيل الدخول بواسطة Google</span>
              </>
            )}
          </button>

          {/* Footer note */}
          <p className="mt-6 text-center text-[12px] leading-relaxed" style={{ color: "#52526a", direction: "rtl" }}>
            لا تملك حساباً؟{" "}
            <a
              href="https://tolzy.me/auth"
              target="_blank"
              rel="noopener noreferrer"
              className="font-bold transition-colors"
              style={{ color: "#818cf8" }}
              onMouseEnter={e => (e.currentTarget.style.color = "#a5b4fc")}
              onMouseLeave={e => (e.currentTarget.style.color = "#818cf8")}
            >
              أنشئ حساباً عبر tolzy.me/auth ←
            </a>
          </p>
        </div>

        {/* Powered-by note */}
        <p className="mt-5 text-center text-[10px]" style={{ color: "#3d3d52" }}>
          مشغّل بواسطة <span className="text-[#52526a] font-semibold">TOLZY Labs</span> · جميع الحقوق محفوظة
        </p>
      </motion.div>
    </main>
  );
}
