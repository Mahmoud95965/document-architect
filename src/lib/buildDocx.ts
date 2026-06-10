import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  ShadingType,
  Footer,
  PageNumber,
} from "docx";
import fileSaver from "file-saver";
const { saveAs } = fileSaver;
import type { DocJSON } from "./generate.functions";
import { isProjectArabic } from "./lang";

const INDIGO = "4F46E5";
const CHARCOAL = "4B5563";
const BODY = "1F2937";
const BORDER = "E5E7EB";
const ZEBRA = "F9FAFB";

const border = { style: BorderStyle.SINGLE, size: 4, color: BORDER };
const cellBorders = { top: border, bottom: border, left: border, right: border };

function buildTable(
  table: { headers: string[]; rows: string[][] },
  isArabic: boolean,
  fontFace: string
) {
  const cols = table.headers.length || 1;
  const totalWidth = 9360; // Standard page printable width in DXA
  const colWidth = Math.floor(totalWidth / cols);
  const columnWidths = Array(cols).fill(colWidth);

  // If RTL (Arabic), reverse columns so they read right-to-left. Otherwise keep original.
  const displayHeaders = isArabic ? [...table.headers].reverse() : table.headers;
  const displayRows = table.rows.map(row => {
    const fullRow = Array.from({ length: cols }, (_, i) => row[i] ?? "");
    return isArabic ? [...fullRow].reverse() : fullRow;
  });

  const headerRow = new TableRow({
    tableHeader: true,
    children: displayHeaders.map(
      (h) =>
        new TableCell({
          borders: cellBorders,
          width: { size: colWidth, type: WidthType.DXA },
          shading: { fill: INDIGO, type: ShadingType.CLEAR, color: "auto" },
          margins: { top: 160, bottom: 160, left: 180, right: 180 }, // Premium padding
          children: [
            new Paragraph({
              alignment: isArabic ? AlignmentType.RIGHT : AlignmentType.LEFT,
              bidirectional: isArabic,
              children: [
                new TextRun({
                  text: h,
                  bold: true,
                  color: "FFFFFF",
                  font: fontFace,
                  size: 20,
                  rightToLeft: isArabic,
                }),
              ],
            }),
          ],
        })
    ),
  });

  const bodyRows = displayRows.map(
    (row, i) =>
      new TableRow({
        children: Array.from({ length: cols }).map((_, c) => {
          const text = row[c] ?? "";
          return new TableCell({
            borders: cellBorders,
            width: { size: colWidth, type: WidthType.DXA },
            shading:
              i % 2 === 1
                ? { fill: ZEBRA, type: ShadingType.CLEAR, color: "auto" }
                : undefined,
            margins: { top: 140, bottom: 140, left: 180, right: 180 },
            children: [
              new Paragraph({
                alignment: isArabic ? AlignmentType.RIGHT : AlignmentType.LEFT,
                bidirectional: isArabic,
                children: [
                  new TextRun({
                    text,
                    font: fontFace,
                    size: 20,
                    color: BODY,
                    rightToLeft: isArabic,
                  }),
                ],
              }),
            ],
          });
        }),
      })
  );

  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths,
    rows: [headerRow, ...bodyRows],
  });
}

export async function generateAndDownloadDocx(data: DocJSON, fileName = "document.docx") {
  const isArabic = isProjectArabic(data);
  const fontFace = isArabic ? "Cairo" : "Segoe UI";
  
  const headingAlignment = isArabic ? AlignmentType.RIGHT : AlignmentType.LEFT;
  const paragraphAlignment = isArabic ? AlignmentType.RIGHT : AlignmentType.JUSTIFIED;

  const children: (Paragraph | Table)[] = [];

  // Elegant top header band instead of raw text title to look premium
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 100, after: 100 },
      children: [
        new TextRun({
          text: "— REPORT & ANALYTICS —",
          bold: true,
          color: INDIGO,
          size: 18,
          font: fontFace,
        }),
      ],
    })
  );

  // Document Title
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 120 },
      bidirectional: isArabic,
      children: [
        new TextRun({
          text: data.title,
          bold: true,
          color: INDIGO,
          size: 44, // 22pt
          font: fontFace,
          rightToLeft: isArabic,
        }),
      ],
    })
  );

  // Subtitle
  if (data.subtitle) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 480 },
        bidirectional: isArabic,
        children: [
          new TextRun({
            text: data.subtitle,
            color: CHARCOAL,
            size: 24, // 12pt
            font: fontFace,
            italics: true,
            rightToLeft: isArabic,
          }),
        ],
      })
    );
  } else {
    children.push(new Paragraph({ spacing: { after: 360 }, children: [] }));
  }

  // Horizontal divider line
  children.push(
    new Paragraph({
      spacing: { after: 360 },
      border: { bottom: { color: BORDER, size: 8, style: BorderStyle.SINGLE } },
      children: [],
    })
  );

  // Sections
  data.sections.forEach((section, idx) => {
    // Heading 1 with elegant vertical accent bar (left in LTR, right in RTL)
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        alignment: headingAlignment,
        bidirectional: isArabic,
        spacing: { before: idx === 0 ? 0 : 400, after: 160 },
        borders: isArabic
          ? { right: { color: INDIGO, size: 24, style: BorderStyle.SINGLE, space: 10 } }
          : { left: { color: INDIGO, size: 24, style: BorderStyle.SINGLE, space: 10 } },
        children: [
          new TextRun({
            text: ` ${section.heading}`,
            bold: true,
            color: INDIGO,
            size: 28, // 14pt
            font: fontFace,
            rightToLeft: isArabic,
          }),
        ],
      })
    );

    // Subheading
    if (section.subheading) {
      children.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          alignment: headingAlignment,
          bidirectional: isArabic,
          spacing: { before: 120, after: 120 },
          children: [
            new TextRun({
              text: section.subheading,
              bold: true,
              color: CHARCOAL,
              size: 24, // 12pt
              font: fontFace,
              rightToLeft: isArabic,
            }),
          ],
        })
      );
    }

    // Paragraphs with proper line spacing and bidirectional settings
    (section.paragraphs ?? []).forEach((p) =>
      children.push(
        new Paragraph({
          alignment: paragraphAlignment,
          bidirectional: isArabic,
          spacing: { after: 160, line: 312 }, // 1.3 line spacing (very readable)
          children: [
            new TextRun({
              text: p,
              color: BODY,
              size: 22, // 11pt
              font: fontFace,
              rightToLeft: isArabic,
            }),
          ],
        })
      )
    );

    // Table
    if (section.table && section.table.headers?.length) {
      children.push(new Paragraph({ spacing: { after: 160 }, children: [] }));
      children.push(buildTable(section.table, isArabic, fontFace));
      children.push(new Paragraph({ spacing: { after: 240 }, children: [] }));
    }
  });

  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: fontFace, size: 22 } },
      },
    },
    numbering: { config: [] },
    sections: [
      {
        properties: {
          page: {
            size: { width: 12240, height: 15840 }, // A4 width/height in dxa
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, // 1 inch margins
          },
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ text: "Page ", color: CHARCOAL, size: 18, font: fontFace }),
                  new TextRun({ children: [PageNumber.CURRENT], color: CHARCOAL, size: 18, font: fontFace }),
                  new TextRun({ text: " of ", color: CHARCOAL, size: 18, font: fontFace }),
                  new TextRun({ children: [PageNumber.TOTAL_PAGES], color: CHARCOAL, size: 18, font: fontFace }),
                ],
              }),
            ],
          }),
        },
        children,
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, fileName);
}
