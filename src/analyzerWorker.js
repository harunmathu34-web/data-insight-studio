import * as XLSX from "xlsx";
import { analyzeWorkbook, workbookInsights } from "./workbookAnalyzer";

self.onmessage = async (event) => {
  try {
    const buffer = event.data;
    self.postMessage({ type: "status", message: "Reading workbook…" });
    const workbook = XLSX.read(buffer, {
      type: "array",
      dense: true,
      cellFormula: false,
      cellHTML: false,
      cellNF: false,
      cellStyles: false,
      cellDates: true,
    });
    self.postMessage({ type: "status", message: `Workbook loaded — ${workbook.SheetNames.length} sheets. Analyzing…` });
    const result = analyzeWorkbook(workbook);
    const insights = workbookInsights(result);
    self.postMessage({ type: "done", result, insights });
  } catch (error) {
    self.postMessage({ type: "error", message: error?.message || String(error) });
  }
};