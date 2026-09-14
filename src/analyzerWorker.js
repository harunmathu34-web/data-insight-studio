import * as XLSX from "xlsx";
import { analyzeWorkbook, workbookInsights } from "./workbookAnalyzer";

self.onmessage = async (event) => {
  try {
    const buffer = event.data;
    self.postMessage({ type: "status", message: "Reading workbook…", progress: 2 });

    const workbook = XLSX.read(buffer, {
      type: "array",
      dense: true,
      cellFormula: false,
      cellHTML: false,
      cellNF: false,
      cellStyles: false,
      cellDates: false,
    });

    const names = workbook.SheetNames || [];
    self.postMessage({
      type: "status",
      message: `Workbook loaded — ${names.length} sheets. Starting analysis…`,
      progress: 5,
    });

    const sheets = [];
    let totalRows = 0;
    let totalColumns = 0;
    let workbookFormulaErrors = 0;
    let workbookDuplicateRows = 0;
    let cohort = null;

    for (let i = 0; i < names.length; i++) {
      const name = names[i];
      const start = 5 + Math.round((i / Math.max(names.length, 1)) * 90);

      self.postMessage({
        type: "status",
        message: `Analyzing ${name} (${i + 1}/${names.length})…`,
        progress: start,
      });

      const oneSheetWorkbook = {
        SheetNames: [name],
        Sheets: { [name]: workbook.Sheets[name] },
      };

      const partial = analyzeWorkbook(oneSheetWorkbook);
      const analyzed = partial.sheets?.[0];

      if (analyzed) {
        sheets.push(analyzed);
        totalRows += analyzed.rows || 0;
        totalColumns += analyzed.columns || 0;
        workbookFormulaErrors += analyzed.quality?.formulaErrors || 0;
        workbookDuplicateRows += analyzed.quality?.duplicateRows || 0;

        if (!cohort && analyzed.cohort) cohort = analyzed.cohort;
        if (/register/i.test(name) && analyzed.cohort) cohort = analyzed.cohort;
      }

      self.postMessage({
        type: "status",
        message: `${name} analyzed — ${analyzed?.rows?.toLocaleString?.() || 0} data rows, ${analyzed?.columns?.toLocaleString?.() || 0} columns.`,
        progress: Math.min(96, start + Math.round(90 / Math.max(names.length, 1))),
      });

      await new Promise(resolve => setTimeout(resolve, 0));
    }

    const result = {
      sheetNames: names,
      sheetCount: sheets.length,
      totalRows,
      totalColumns,
      sheets,
      cohort,
      workbookQuality: {
        formulaErrors: workbookFormulaErrors,
        duplicateRows: workbookDuplicateRows,
      },
    };

    const insights = workbookInsights(result);
    self.postMessage({ type: "done", result, insights, progress: 100 });
  } catch (error) {
    self.postMessage({ type: "error", message: error?.stack || error?.message || String(error) });
  }
};
