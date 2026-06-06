const fs = require('fs');
let code = fs.readFileSync('src/routes/index.tsx', 'utf8');

// Inject isDark state into Index() component
if (!code.includes('const [isDark, setIsDark]')) {
  code = code.replace(
    /function Index\(\) \{/,
    `function Index() {
  const [isDark, setIsDark] = useState(true);
  
  useEffect(() => {
    if (isDark) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  }, [isDark]);`
  );
}

// Fix 'e' implicit any types
code = code.replace(/onMouseEnter=\{e =>/g, 'onMouseEnter={(e: any) =>');
code = code.replace(/onMouseLeave=\{e =>/g, 'onMouseLeave={(e: any) =>');
code = code.replace(/onClick=\{e =>/g, 'onClick={(e: any) =>');

fs.writeFileSync('src/routes/index.tsx', code);
console.log("Fixed index.tsx successfully");
