import ExcelJS from "exceljs";
import fileSaver from "file-saver";
const { saveAs } = fileSaver;
import type { ExcelJSON } from "./generate.functions";
import { isProjectArabic } from "./lang";

export async function generateAndDownloadExcel(data: ExcelJSON, fileName = "spreadsheet.xlsx") {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Docify";
  workbook.lastModifiedBy = "Docify";
  workbook.created = new Date();
  workbook.modified = new Date();

  const isArabic = isProjectArabic(data);
  const textAlignment = isArabic ? "right" : "left";

  data.sheets.forEach((sheetData) => {
    // Add worksheet with dynamic RTL view based on language detection
    const worksheet = workbook.addWorksheet(sheetData.name, {
      views: [{ showGridLines: true, rightToLeft: isArabic }]
    });

    // Add headers
    if (sheetData.headers && sheetData.headers.length > 0) {
      const headerRow = worksheet.addRow(sheetData.headers);
      headerRow.font = { name: "Arial", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      headerRow.alignment = { vertical: "middle", horizontal: textAlignment };
      headerRow.height = 26;

      headerRow.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF1E3A8A" } // NAVY header background
        };
        cell.border = {
          top: { style: "thin", color: { argb: "FFE5E7EB" } },
          bottom: { style: "thin", color: { argb: "FFE5E7EB" } },
          left: { style: "thin", color: { argb: "FFE5E7EB" } },
          right: { style: "thin", color: { argb: "FFE5E7EB" } }
        };
      });
    }

    // Add rows
    if (sheetData.rows && sheetData.rows.length > 0) {
      sheetData.rows.forEach((rowCells, rIdx) => {
        // Map ExcelCell array to ExcelJS row values
        const rowValues = rowCells.map((c) => {
          if (c.formula) {
            // Strip leading '=' if present, ExcelJS expects formula without '='
            return { formula: c.formula.startsWith("=") ? c.formula.substring(1) : c.formula };
          }
          return c.value;
        });

        const addedRow = worksheet.addRow(rowValues);
        addedRow.height = 21;

        const isEvenRow = rIdx % 2 === 0;

        rowCells.forEach((c, cIdx) => {
          const cell = addedRow.getCell(cIdx + 1);
          
          // Determine if value is numeric to right-align numbers globally
          const isNumeric = 
            typeof c.value === "number" || 
            (!isNaN(Number(c.value)) && String(c.value).trim() !== "" && !c.value.toString().startsWith("0"));

          cell.font = {
            name: "Arial",
            size: 10,
            bold: c.bold ?? false,
            color: c.color ? { argb: "FF" + c.color.replace("#", "") } : { argb: "FF374151" }
          };

          cell.alignment = {
            vertical: "middle",
            horizontal: isNumeric ? "right" : textAlignment,
          };

          // Apply bg pattern if defined, otherwise apply modern light zebra striping
          const bgHex = c.bg ? c.bg.replace("#", "") : (isEvenRow ? "F9FAFB" : null);
          if (bgHex) {
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FF" + bgHex }
            };
          }

          // Apply borders
          cell.border = {
            top: { style: "thin", color: { argb: "FFE5E7EB" } },
            bottom: { style: "thin", color: { argb: "FFE5E7EB" } },
            left: { style: "thin", color: { argb: "FFE5E7EB" } },
            right: { style: "thin", color: { argb: "FFE5E7EB" } }
          };
        });
      });
    }

    // Auto-fit column widths
    worksheet.columns.forEach((column) => {
      let maxLen = 12;
      column.eachCell!({ includeEmpty: true }, (cell) => {
        let valStr = "";
        if (cell.value) {
          if (typeof cell.value === "object" && "formula" in cell.value) {
            valStr = String(cell.value.formula || "");
          } else {
            valStr = String(cell.value);
          }
        }
        if (valStr.length > maxLen) {
          maxLen = valStr.length;
        }
      });
      column.width = Math.min(maxLen + 4, 45); // Set maximum safe column width limit
    });
  });

  // Write workbook to buffer and save
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  saveAs(blob, fileName);
}
