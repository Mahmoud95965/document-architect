import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import {
  Outlet,
  Link,
  createRootRouteWithContext,
  useRouter,
  HeadContent,
  Scripts,
} from "@tanstack/react-router";

import appCss from "../styles.css?url";
import { AuthProvider } from "@/hooks/useAuth";

function NotFoundComponent() {
  return (
    <div className="flex min-h-screen items-center justify-center bg-[#050507] px-4">
      <div className="max-w-md text-center animate-float-up">
        <div className="mx-auto mb-6 w-20 h-20 rounded-3xl glass-card flex items-center justify-center">
          <span className="text-4xl">🔍</span>
        </div>
        <h1 className="text-8xl font-black text-gradient">404</h1>
        <h2 className="mt-4 text-xl font-bold text-white">الصفحة غير موجودة</h2>
        <p className="mt-2 text-sm text-[#6b6b7b] leading-relaxed">
          الصفحة التي تبحث عنها غير موجودة أو تم نقلها.
        </p>
        <div className="mt-8">
          <Link
            to="/"
            className="inline-flex items-center justify-center gap-2 rounded-2xl bg-indigo-600 hover:bg-indigo-500 px-6 py-3 text-sm font-semibold text-white transition-all duration-200 shadow-lg shadow-indigo-900/30 btn-shimmer"
          >
            العودة للرئيسية
          </Link>
        </div>
      </div>
    </div>
  );
}

function ErrorComponent({ error, reset }: { error: Error; reset: () => void }) {
  console.error(error);
  const router = useRouter();

  return (
    <div className="flex min-h-screen items-center justify-center bg-[#050507] px-4">
      <div className="max-w-md text-center animate-float-up">
        <div className="mx-auto mb-6 w-20 h-20 rounded-3xl glass-card flex items-center justify-center">
          <span className="text-4xl">⚠️</span>
        </div>
        <h1 className="text-2xl font-bold text-white">حدث خطأ ما</h1>
        <p className="mt-3 text-sm text-[#6b6b7b] leading-relaxed">
          حدث خطأ غير متوقع. يمكنك المحاولة مرة أخرى أو العودة للصفحة الرئيسية.
        </p>
        <div className="mt-8 flex flex-wrap justify-center gap-3">
          <button
            onClick={() => {
              router.invalidate();
              reset();
            }}
            className="inline-flex items-center gap-2 rounded-2xl bg-indigo-600 hover:bg-indigo-500 px-5 py-2.5 text-sm font-semibold text-white transition-all btn-shimmer"
          >
            حاول مجدداً
          </button>
          <a
            href="/"
            className="inline-flex items-center gap-2 rounded-2xl glass border border-white/10 hover:bg-white/5 px-5 py-2.5 text-sm font-medium text-white transition-all"
          >
            الرئيسية
          </a>
        </div>
      </div>
    </div>
  );
}

export const Route = createRootRouteWithContext<{ queryClient: QueryClient }>()({
  head: () => ({
    meta: [
      { charSet: "utf-8" },
      { name: "viewport", content: "width=device-width, initial-scale=1" },
      { title: "AXIOM Flow — توليد وتصميم المستندات والشرائح بالذكاء الاصطناعي (Word, Excel, PPT)" },
      {
        name: "description",
        content:
          "حوّل أفكارك إلى مستندات Word احترافية، جداول Excel محاسبية، وعروض PowerPoint سينمائية تفاعلية في ثوانٍ بقوة محرك الذكاء الاصطناعي AXIOM Flow.",
      },
      {
        name: "keywords",
        content: "AXIOM Flow, توليد مستندات بالذكاء الاصطناعي, تصميم بوربوينت تلقائي, إكسل محاسبي بالذكاء الاصطناعي, وورد عربي بالذكاء الاصطناعي, AI document generator, Word Excel PowerPoint AI creator",
      },
      { name: "author", content: "AXIOM Labs" },
      { property: "og:title", content: "AXIOM Flow — AI Word, Excel & PowerPoint Generator" },
      { property: "og:description", content: "قم بتصميم وتوليد ملفات Word وExcel وPowerPoint تلقائياً بالذكاء الاصطناعي مع دعم كامل للغة العربية والتنسيقات المحاسبية الرسمية." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
      { name: "twitter:site", content: "@AXIOMLabs" },
      { name: "theme-color", content: "#050507" },
    ],
    links: [
      { rel: "preconnect", href: "https://fonts.googleapis.com" },
      { rel: "preconnect", href: "https://fonts.gstatic.com", crossOrigin: "anonymous" as const },
      {
        rel: "stylesheet",
        href: "https://fonts.googleapis.com/css2?family=Cairo:wght@200..1000&family=Inter:wght@300;400;500;600;700;800;900&family=JetBrains+Mono:wght@400;500;600&display=swap",
      },
      { rel: "stylesheet", href: appCss },
      { rel: "icon", type: "image/jpeg", href: "/logo.jpg" },
      { rel: "apple-touch-icon", href: "/logo.jpg" },
    ],
  }),
  shellComponent: RootShell,
  component: RootComponent,
  notFoundComponent: NotFoundComponent,
  errorComponent: ErrorComponent,
});

function RootShell({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ar" dir="rtl">
      <head>
        <HeadContent />
      </head>
      <body>
        {children}
        <Scripts />
      </body>
    </html>
  );
}

function RootComponent() {
  const { queryClient } = Route.useRouteContext();

  return (
    <QueryClientProvider client={queryClient}>
      {/* AuthProvider wraps everything so all routes have access to auth state */}
      <AuthProvider>
        <Outlet />
      </AuthProvider>
    </QueryClientProvider>
  );
}
