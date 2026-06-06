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

const FONT = "Calibri";
const NAVY = "1E3A8A";
const CHARCOAL = "374151";
const BODY = "1F2937";
const BORDER = "E5E7EB";
const ZEBRA = "F9FAFB";

const border = { style: BorderStyle.SINGLE, size: 4, color: BORDER };
const cellBorders = { top: border, bottom: border, left: border, right: border };

function buildTable(table: { headers: string[]; rows: string[][] }, isArabic: boolean) {
  const cols = table.headers.length || 1;
  const totalWidth = 9360;
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
          shading: { fill: NAVY, type: ShadingType.CLEAR, color: "auto" },
          margins: { top: 140, bottom: 140, left: 160, right: 160 },
          children: [
            new Paragraph({
              alignment: isArabic ? AlignmentType.RIGHT : AlignmentType.LEFT,
              children: [
                new TextRun({ text: h, bold: true, color: "FFFFFF", font: FONT, size: 22, }),
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
            margins: { top: 120, bottom: 120, left: 160, right: 160 },
            children: [
              new Paragraph({
                alignment: isArabic ? AlignmentType.RIGHT : AlignmentType.LEFT,
                children: [new TextRun({ text, font: FONT, size: 22, color: BODY, })],
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
  const headingAlignment = isArabic ? AlignmentType.RIGHT : AlignmentType.LEFT;
  const paragraphAlignment = isArabic ? AlignmentType.RIGHT : AlignmentType.JUSTIFIED;

  const children: (Paragraph | Table)[] = [];

  // Document Title
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 120 },
      children: [
        new TextRun({ text: data.title, bold: true, color: NAVY, size: 48, font: FONT, }),
      ],
    })
  );

  // Subtitle
  if (data.subtitle) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 480 },
        children: [
          new TextRun({ text: data.subtitle, color: CHARCOAL, size: 24, font: FONT, italics: true, }),
        ],
      })
    );
  } else {
    children.push(new Paragraph({ spacing: { after: 360 }, children: [] }));
  }

  // Sections
  data.sections.forEach((section, idx) => {
    // Heading 1
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        alignment: headingAlignment,
        spacing: { before: idx === 0 ? 0 : 360, after: 160 },
        children: [
          new TextRun({ text: section.heading, bold: true, color: NAVY, size: 32, font: FONT, }),
        ],
      })
    );

    // Subheading
    if (section.subheading) {
      children.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          alignment: headingAlignment,
          spacing: { before: 80, after: 120 },
          children: [
            new TextRun({ text: section.subheading, bold: true, color: CHARCOAL, size: 26, font: FONT, }),
          ],
        })
      );
    }

    // Paragraphs
    (section.paragraphs ?? []).forEach((p) =>
      children.push(
        new Paragraph({
          alignment: paragraphAlignment,
          spacing: { after: 160, line: 276 }, // 1.15 line spacing
          children: [new TextRun({ text: p, color: BODY, size: 23, font: FONT, })],
        })
      )
    );

    // Table
    if (section.table && section.table.headers?.length) {
      children.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
      children.push(buildTable(section.table, isArabic));
      children.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
    }
  });

  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: FONT, size: 23 } },
      },
    },
    numbering: { config: [] },
    sections: [
      {
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ text: "Page ", color: CHARCOAL, size: 18, font: FONT }),
                  new TextRun({ children: [PageNumber.CURRENT], color: CHARCOAL, size: 18, font: FONT }),
                  new TextRun({ text: " of ", color: CHARCOAL, size: 18, font: FONT }),
                  new TextRun({ children: [PageNumber.TOTAL_PAGES], color: CHARCOAL, size: 18, font: FONT }),
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
