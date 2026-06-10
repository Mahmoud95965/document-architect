import ExcelJS from "exceljs";
import fileSaver from "file-saver";
const { saveAs } = fileSaver;
import type { ExcelJSON } from "./generate.functions";
import { isProjectArabic } from "./lang";

export async function generateAndDownloadExcel(data: ExcelJSON, fileName = "spreadsheet.xlsx") {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "TOLZY Flow";
  workbook.lastModifiedBy = "TOLZY Flow";
  workbook.created = new Date();
  workbook.modified = new Date();

  const isArabic = isProjectArabic(data);
  const textAlignment = isArabic ? "right" : "left";
  const fontName = isArabic ? "Cairo" : "Segoe UI";

  data.sheets.forEach((sheetData) => {
    // Add worksheet with dynamic RTL view based on language detection
    const worksheet = workbook.addWorksheet(sheetData.name, {
      views: [{ showGridLines: true, rightToLeft: isArabic }]
    });

    // Add headers
    if (sheetData.headers && sheetData.headers.length > 0) {
      const headerRow = worksheet.addRow(sheetData.headers);
      headerRow.font = { name: fontName, size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      headerRow.alignment = { vertical: "middle", horizontal: textAlignment };
      headerRow.height = 28; // Premium row height

      headerRow.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF107C41" } // Excel green theme
        };
        cell.border = {
          top: { style: "thin", color: { argb: "FF0B5E31" } },
          bottom: { style: "medium", color: { argb: "FF0B5E31" } },
          left: { style: "thin", color: { argb: "FF0B5E31" } },
          right: { style: "thin", color: { argb: "FF0B5E31" } }
        };
      });
    }

    // Add rows
    if (sheetData.rows && sheetData.rows.length > 0) {
      sheetData.rows.forEach((rowCells, rIdx) => {
        // Detect if this row represents a total/summary row
        const isTotalRow = rowCells.some((cellObj) => {
          const valStr = String(cellObj.value || "").toLowerCase();
          return (
            valStr.includes("إجمالي") ||
            valStr.includes("المجموع") ||
            valStr.includes("مجموع") ||
            valStr.includes("total") ||
            valStr.includes("average") ||
            valStr.includes("المتوسط")
          );
        });

        // Map ExcelCell array to ExcelJS row values
        const rowValues = rowCells.map((c) => {
          if (c.formula) {
            // Strip leading '=' if present, ExcelJS expects formula without '='
            return { formula: c.formula.startsWith("=") ? c.formula.substring(1) : c.formula };
          }
          
          // Force numbers to be stored as actual numbers instead of strings for calculations
          const isNumericStr = 
            typeof c.value === "string" && 
            !isNaN(Number(c.value)) && 
            c.value.trim() !== "" && 
            !c.value.startsWith("0");
            
          return isNumericStr ? Number(c.value) : c.value;
        });

        const addedRow = worksheet.addRow(rowValues);
        addedRow.height = isTotalRow ? 24 : 22; // Premium spacing

        const isEvenRow = rIdx % 2 === 0;

        rowCells.forEach((c, cIdx) => {
          const cell = addedRow.getCell(cIdx + 1);
          
          // Determine if value is numeric to right-align and format numbers globally
          const isNumericVal = 
            typeof cell.value === "number" || 
            (typeof cell.value === "object" && cell.value !== null && "formula" in cell.value);

          cell.font = {
            name: fontName,
            size: isTotalRow ? 10.5 : 10,
            bold: isTotalRow ? true : (c.bold ?? false),
            color: c.color ? { argb: "FF" + c.color.replace("#", "") } : { argb: isTotalRow ? "FF0B5E31" : "FF374151" }
          };

          cell.alignment = {
            vertical: "middle",
            horizontal: isNumericVal ? "right" : textAlignment,
          };

          // Apply number formatting for numeric cells (comma separators)
          if (isNumericVal) {
            const numValue = Number(cell.value);
            // Check if there are decimals
            if (!isNaN(numValue)) {
              cell.numFmt = numValue % 1 !== 0 ? "#,##0.00" : "#,##0";
            } else {
              // If it's a formula, apply general numeric format with commas
              cell.numFmt = "#,##0.00";
            }
          }

          // Background styling: Total rows get a soft green tint, others alternate zebra striping
          let bgHex = c.bg ? c.bg.replace("#", "") : null;
          if (!bgHex) {
            if (isTotalRow) {
              bgHex = "FFE6F4EA"; // accounting green tint
            } else if (isEvenRow) {
              bgHex = "FFF4FBF8"; // ultra-soft green alternating striping
            }
          }

          if (bgHex) {
            // If it starts with FF, keep it, otherwise prepend FF for alpha channel
            const argbColor = bgHex.length === 8 ? bgHex : "FF" + bgHex;
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: argbColor }
            };
          }

          // Borders: Total row gets thin top and double bottom border (Accounting Double Underline)
          if (isTotalRow) {
            cell.border = {
              top: { style: "thin", color: { argb: "FFCCCCCC" } },
              bottom: { style: "double", color: { argb: "FF107C41" } }, // Double green accounting border
              left: { style: "thin", color: { argb: "FFE5E7EB" } },
              right: { style: "thin", color: { argb: "FFE5E7EB" } }
            };
          } else {
            cell.border = {
              top: { style: "thin", color: { argb: "FFE5E7EB" } },
              bottom: { style: "thin", color: { argb: "FFE5E7EB" } },
              left: { style: "thin", color: { argb: "FFE5E7EB" } },
              right: { style: "thin", color: { argb: "FFE5E7EB" } }
            };
          }
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
