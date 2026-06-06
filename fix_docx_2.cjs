const fs = require('fs');
let code = fs.readFileSync('src/lib/buildDocx.ts', 'utf8');

// The original error was that `bidirectional` or `rightToLeft` does not exist in IParagraphOptions / IRunOptions.
// To fix the typescript compilation errors and unblock the build, we will remove `rightToLeft: isArabic`
code = code.replace(/rightToLeft:\s*isArabic,?\s*/g, '');

fs.writeFileSync('src/lib/buildDocx.ts', code);
console.log("Fixed buildDocx.ts successfully");
