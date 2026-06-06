const fs = require('fs');
let code = fs.readFileSync('src/routes/index.tsx', 'utf8');

// 1. Add isDark state and Moon/Sun to lucide-react imports
code = code.replace(/import {([^}]+)} from "lucide-react";/, (match, p1) => {
  if (!p1.includes('Moon')) {
    return `import {${p1}, Moon, Sun} from "lucide-react";`;
  }
  return match;
});

// 2. Add isDark state
if (!code.includes('const [isDark, setIsDark]')) {
  code = code.replace(
    /export default function Home\(\) \{/,
    `export default function Home() {
  const [isDark, setIsDark] = React.useState(true);
  
  React.useEffect(() => {
    if (isDark) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  }, [isDark]);`
  );
}

// 3. Add Theme Toggle to Header Nav
const themeToggleBtn = `
          {/* Theme Toggle */}
          <button
            onClick={() => setIsDark(!isDark)}
            className="flex h-9 w-9 items-center justify-center rounded-full transition-all border border-black/5 dark:border-white/10 text-slate-500 hover:text-slate-900 dark:text-slate-300 dark:hover:text-white"
            style={{
              background: isDark ? "rgba(255,255,255,0.04)" : "rgba(0,0,0,0.02)",
            }}
          >
            {isDark ? <Sun className="h-4 w-4" /> : <Moon className="h-4 w-4" />}
          </button>
`;
if (!code.includes('Theme Toggle')) {
  code = code.replace(
    /\{\/\* User Profile Menu \*\/\}/,
    `${themeToggleBtn}\n          {/* User Profile Menu */}`
  );
}

// 4. Color replacements
code = code.replace(/bg-\[\#050507\]/g, 'bg-background');
code = code.replace(/text-\[\#6b6b7b\]/g, 'text-muted-foreground');
code = code.replace(/text-\[\#52526a\]/g, 'text-muted-foreground');
code = code.replace(/text-\[\#4b4b5b\]/g, 'text-muted-foreground');
code = code.replace(/border-white\/10/g, 'border-border');
code = code.replace(/border-white\/5/g, 'border-border/50');

// 5. Arabic Translations
code = code.replace(/"Automated Spreadsheet"/g, '"جداول إكسل ذكية"');
code = code.replace(/"Stunning Slide Decks"/g, '"عروض تقديمية مبهرة"');
code = code.replace(/"Professional Documents"/g, '"مستندات احترافية"');
code = code.replace(/"Built by AI."/g, '"مدعومة بالذكاء الاصطناعي."');
code = code.replace(/Upgrade to Ultra/g, 'ترقية إلى Ultra');
code = code.replace(/AI Document Architect/g, '');
code = code.replace(/>Recent Projects</g, '>المشاريع الأخيرة<');
code = code.replace(/placeholder="Describe the document you want to create\.\.\. "/g, 'placeholder="صف المستند الذي تريد إنشاؤه بدقة..."');
code = code.replace(/placeholder="Enter the main title\.\.\."/g, 'placeholder="أدخل العنوان الرئيسي..."');
code = code.replace(/placeholder="e\.g\., Financial Report 2024"/g, 'placeholder="مثال: التقرير المالي لعام 2024"');
code = code.replace(/>Create Document</g, '>إنشاء المستند<');
code = code.replace(/>View All</g, '>عرض الكل<');
code = code.replace(/>Create New</g, '>مشروع جديد<');
code = code.replace(/Generate Content/g, 'توليد المحتوى');
code = code.replace(/Preparing your project\.\.\./g, 'جاري تحضير مشروعك...');
code = code.replace(/>Download</g, '>تحميل<');
code = code.replace(/>Back</g, '>رجوع<');
code = code.replace(/Search projects\.\.\./g, 'ابحث في المشاريع...');

// 6. Projects Grid Redesign
// Find the projects list container and wrap the mapping in a grid
code = code.replace(
  /<div className="mt-4 flex flex-col gap-2">([\s\S]*?)<\/div>\s*<\/div>\s*<\/motion\.div>/,
  `<div className="mt-4 grid grid-cols-1 md:grid-cols-2 gap-3">$1</div>
            </div>
          </motion.div>`
);

// 7. Update glass card styling in project mapping for better dark/light support
code = code.replace(
  /className="group\/item flex items-center justify-between p-3 rounded-xl border cursor-pointer transition-all"/g,
  'className="group/item flex items-center justify-between p-4 rounded-xl border cursor-pointer transition-all bg-card border-border hover:shadow-md"'
);
code = code.replace(
  /style=\{\{\s*background: "rgba\(255,255,255,0\.02\)",\s*borderColor: "rgba\(255,255,255,0\.065\)",\s*\}\}/g,
  ''
);

// Removing inline hover styles for projects as tailwind handles it
code = code.replace(/onMouseEnter=\{e => \{\s*\(e\.currentTarget as HTMLElement\)\.style\.background = "rgba\(255,255,255,0\.04\)";\s*\(e\.currentTarget as HTMLElement\)\.style\.borderColor = "rgba\(255,255,255,0\.12\)";\s*\}\}/g, '');
code = code.replace(/onMouseLeave=\{e => \{\s*\(e\.currentTarget as HTMLElement\)\.style\.background = "rgba\(255,255,255,0\.02\)";\s*\(e\.currentTarget as HTMLElement\)\.style\.borderColor = "rgba\(255,255,255,0\.065\)";\s*\}\}/g, '');

fs.writeFileSync('src/routes/index.tsx', code);
console.log("Updated index.tsx successfully");
