export function isArabicText(text = ""): boolean {
  const arabicRegex = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;
  return arabicRegex.test(text);
}

export function isProjectArabic(project: {
  title: string;
  prompt?: string;
  sections?: any[];
  sheets?: any[];
  slides?: any[];
}): boolean {
  if (isArabicText(project.title)) return true;
  if (project.prompt && isArabicText(project.prompt)) return true;

  // Scan sections (Word)
  if (project.sections && project.sections.length > 0) {
    for (const section of project.sections) {
      if (isArabicText(section.heading) || isArabicText(section.subheading || "")) return true;
      if (section.paragraphs && section.paragraphs.some((p: string) => isArabicText(p))) return true;
    }
  }

  // Scan sheets (Excel)
  if (project.sheets && project.sheets.length > 0) {
    for (const sheet of project.sheets) {
      if (isArabicText(sheet.name)) return true;
      if (sheet.headers && sheet.headers.some((h: string) => isArabicText(h))) return true;
      if (sheet.rows && sheet.rows.some((row: any[]) => row.some((cell: any) => isArabicText(String(cell?.value || ""))))) return true;
    }
  }

  // Scan slides (PowerPoint)
  if (project.slides && project.slides.length > 0) {
    for (const slide of project.slides) {
      if (isArabicText(slide.title)) return true;
      if (slide.bullets && slide.bullets.some((b: string) => isArabicText(b))) return true;
    }
  }

  return false;
}
