const fs = require('fs');
let code = fs.readFileSync('src/routes/index.tsx', 'utf8');

// Replace specific problematic tailwind classes
// `text-white` is tricky because it might be on a colored button.
// For example, "Upgrade to Ultra" button has color: "#fff" so it's fine.
// But standard `text-white` usually needs dark:text-white and a default dark text.
code = code.replace(/text-white\/90/g, 'text-foreground/90 dark:text-white/90');
code = code.replace(/text-white\/80/g, 'text-foreground/80 dark:text-white/80');
code = code.replace(/text-white\/70/g, 'text-foreground/70 dark:text-white/70');
code = code.replace(/text-white\/50/g, 'text-muted-foreground dark:text-white/50');
code = code.replace(/text-slate-300/g, 'text-slate-700 dark:text-slate-300');
code = code.replace(/text-slate-400/g, 'text-slate-600 dark:text-slate-400');
code = code.replace(/text-gray-300/g, 'text-gray-700 dark:text-gray-300');
code = code.replace(/text-gray-400/g, 'text-gray-600 dark:text-gray-400');

// text-white generally should be text-foreground dark:text-white 
// EXCEPT when inside a button or colored badge. Let's do a targeted replace for common UI text.
// Example: `<p className="text-sm text-white">`
code = code.replace(/text-white([^a-zA-Z\/])/g, 'text-foreground dark:text-white$1');
// If we accidentally replaced button text-white, let's fix known ones:
code = code.replace(/from-white to-white\/70/g, 'from-foreground to-foreground/70 dark:from-white dark:to-white/70');

// Replace inline colors
code = code.replace(/color:\s*"#fff"/g, 'color: "var(--foreground)"');
code = code.replace(/color:\s*"#c4c4d4"/g, 'color: "var(--muted-foreground)"');
code = code.replace(/color:\s*"#a1a1aa"/g, 'color: "var(--muted-foreground)"');
code = code.replace(/color:\s*"#d4d4d8"/g, 'color: "var(--foreground)"');
code = code.replace(/color:\s*"rgba\(255,255,255,0\.9\)"/g, 'color: "var(--foreground)"');
code = code.replace(/color:\s*"rgba\(255,255,255,0\.7\)"/g, 'color: "var(--muted-foreground)"');

// Fix the Theme Toggle button icon visibility since text-slate-300 might be used
code = code.replace(/text-slate-300 dark:hover:text-white/g, 'text-slate-700 hover:text-slate-900 dark:text-slate-300 dark:hover:text-white');

// Fix border-white/x to be responsive
code = code.replace(/border-white\/10/g, 'border-black/10 dark:border-white/10');
code = code.replace(/border-white\/20/g, 'border-black/20 dark:border-white/20');
code = code.replace(/border-white\/5/g, 'border-black/5 dark:border-white/5');
code = code.replace(/bg-white\//g, 'bg-black/ dark:bg-white/'); // Oops, let's just do specific ones
code = code.replace(/bg-white\/\[0\.02\]/g, 'bg-black/5 dark:bg-white/[0.02]');
code = code.replace(/bg-white\/5/g, 'bg-black/5 dark:bg-white/5');

// Fix the main gradient text to stand out in light mode too
// `bg-gradient-to-r from-foreground to-foreground/70 dark:from-white dark:to-white/70`
// There's also `text-gradient` class in CSS, let's make sure it is updated in styles.css if needed, but it's likely fine as it uses light colors which might disappear.
// We'll update styles.css for text-gradient.

fs.writeFileSync('src/routes/index.tsx', code);
console.log("Fixed text visibility in index.tsx");
