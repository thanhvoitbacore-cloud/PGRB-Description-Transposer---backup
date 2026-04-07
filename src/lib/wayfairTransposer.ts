import * as XLSX from "xlsx";

export interface TransposedRow {
  sku: string;
  base_heading: string;
  attribute_heading: string;
  value: string;
  validation_status: "Valid" | "Duplicate_Key" | "Corrupted_Value";
}

const EXCEL_ERRORS = ["#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NULL!", "#NAME?", "#NUM!"];

export function validateFileName(name: string): boolean {
  return name.toLowerCase().includes("product description export");
}

function cleanValue(val: unknown): string {
  if (val === null || val === undefined) return "";
  return String(val).replace(/[\r\n]+/g, " ").trim();
}

function isBlank(val: string): boolean {
  return val === "" || val.trim() === "";
}

export function processWorkbook(workbook: XLSX.WorkBook): TransposedRow[] {
  const sheetNames = workbook.SheetNames;
  if (sheetNames.length === 0) throw new Error("File không có sheet nào.");

  const normalize = (value: string) => value.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();

  const findHeaderRowIndex = (raw: unknown[][]): number => {
    for (let i = 0; i < Math.min(raw.length, 10); i++) {
      const cells = (raw[i] || []).map((cell) => cleanValue(cell));
      const normalized = cells.map(normalize);
      const hasSku = normalized.some((cell) => cell === "sku" || cell.includes("supplier part") || cell.includes("partner sku"));
      const hasFeature = normalized.some((cell) => cell.includes("feature bullet"));
      if (hasSku && (hasFeature || cells.filter(Boolean).length >= 3)) return i;
    }
    return 3;
  };

  const findSkuColumn = (headers: string[]): string | undefined => {
    return headers.find((header) => {
      const value = normalize(header);
      return (
        value.includes("wayfair listing") ||
        value === "sku" ||
        value.includes("supplier part") ||
        value.includes("partner sku") ||
        value.includes("part #") ||
        value.includes("part number") ||
        value.includes("wayfair sku") ||
        value.includes("wayfair part")
      );
    });
  };

  const findFilterColumn = (headers: string[]): string | undefined => {
    return headers.find((header) => {
      const value = normalize(header);
      return value.includes("manufacturer part number");
    });
  };

  const isTemplateHelperRow = (row: Record<string, string>): boolean => {
    const helperKeywords = [
      "text", 
      "drop down", 
      "character limit", 
      "maximum", 
      "minimum", 
      "required", 
      "optional", 
      "format",
      "instruction",
    ];
    
    const nonEmptyValues = Object.values(row).filter((v) => !isBlank(v));
    if (nonEmptyValues.length === 0) return true;
    
    const helperCount = nonEmptyValues.filter((v) => {
      const low = v.toLowerCase();
      return helperKeywords.some((kw) => low.includes(kw));
    }).length;
    
    return helperCount / nonEmptyValues.length > 0.3;
  };

  const parseSheet = (name: string, index: number): { headers: string[]; rows: Record<string, string>[]; skuCol?: string; filterCol?: string; sheetName: string; sheetIndex: number } => {
    const ws = workbook.Sheets[name];
    
    // Dynamically recalculate !ref to prevent truncated data
    let maxRow = 0;
    let maxCol = 0;
    for (const key in ws) {
      if (key.startsWith("!")) continue;
      const cell = XLSX.utils.decode_cell(key);
      if (cell.r > maxRow) maxRow = cell.r;
      if (cell.c > maxCol) maxCol = cell.c;
    }
    ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: maxCol } });

    const raw: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: false });
    if (raw.length === 0) return { headers: [], rows: [], sheetName: name, sheetIndex: index };

    const headerRowIndex = findHeaderRowIndex(raw);
    const sourceHeaders = raw[headerRowIndex] || [];
    const headers = sourceHeaders.map((h, idx) => cleanValue(h) || `Column ${idx + 1}`);
    const skuCol = findSkuColumn(headers);
    const filterCol = findFilterColumn(headers);
    const rows: Record<string, string>[] = [];

    for (let i = headerRowIndex + 1; i < raw.length; i++) {
      const values = raw[i] || [];
      if (values.every((cell) => isBlank(cleanValue(cell)))) continue;
      
      const row: Record<string, string> = {
        _excel_row: String(i + 1),
      };
      headers.forEach((h, idx) => {
        row[h] = cleanValue(values[idx]);
      });

      if (isTemplateHelperRow(row)) continue;
      
      rows.push(row);
    }

    return { headers, rows, skuCol, filterCol, sheetName: name, sheetIndex: index };
  };

  const parsedSheets = sheetNames.map((name, idx) => parseSheet(name, idx + 1));
  
  const scoredSheets = parsedSheets.map((sheet) => {
    let score = 0;
    if (sheet.skuCol) score += 100;
    if (sheet.headers.length > 5) score += sheet.headers.length;
    if (sheet.headers.some((h) => normalize(h).includes("feature bullet"))) score += 50;
    if (sheet.rows.length > 0) score += 10;
    
    const lowName = sheet.sheetName.toLowerCase();
    
    const sheetVisibility = (workbook as any).Workbook?.Sheets?.find((s: any) => s.name === sheet.sheetName);
    const isVisible = !sheetVisibility || sheetVisibility.Hidden === 0;

    if (!isVisible) {
      score -= 2000;
    } else {
      if (lowName === "products" || lowName === "wayfair" || lowName.includes("product description")) {
        score += 500;
      }
    }

    if (/\(\d+\)$/.test(sheet.sheetName) || sheet.sheetName.toLowerCase().includes("copy")) {
      score -= 1500;
    }

    if (lowName.includes("instruction") || lowName.includes("intro") || lowName.includes("summary")) {
      score -= 500;
    }
    
    return { ...sheet, score, isVisible };
  });

  const main = scoredSheets.reduce((prev, current) => (prev.score > current.score ? prev : current));
  
  if (!main.skuCol || main.score < 50) {
    throw new Error("Không tìm thấy cột SKU (Wayfair Listing) trong bất kỳ sheet nào.");
  }

  const bulletRegex = /((additional\s*)?feature\s*(bullet)?\s*(\d+)?)|(bullet\s*\d+)/i;
  const marketingCol = main.headers.find((h) => normalize(h).includes("marketing copy"));
  const featureCols = main.headers.filter((h) => bulletRegex.test(h));

  const mainData = new Map<string, { attrs: Record<string, string>, rowNum: string }>();
  for (const row of main.rows) {
    const sku = row[main.skuCol!];
    const filterValue = main.filterCol ? row[main.filterCol] : "";
    
    if (isBlank(sku)) continue;
    
    if (main.filterCol && normalize(filterValue) !== "wayfair sku") {
      continue;
    }

    const attrs: Record<string, string> = {};
    
    if (marketingCol) {
      attrs[marketingCol] = row[marketingCol] || "";
    }

    for (const col of featureCols) {
      const value = row[col] || "";
      if (!isBlank(value)) attrs[col] = value;
    }

    mainData.set(sku, { attrs, rowNum: row._excel_row });
  }

  // Fallback pass: some products in the export don't have a row explicitly labeled
  // "Wayfair SKU" in the Manufacturer Part Number column, but they still have valid
  // Wayfair Listing values in their other rows. Capture any SKUs not yet in mainData.
  if (main.filterCol) {
    for (const row of main.rows) {
      const sku = row[main.skuCol!];
      if (isBlank(sku) || mainData.has(sku)) continue;

      const attrs: Record<string, string> = {};

      if (marketingCol) {
        attrs[marketingCol] = row[marketingCol] || "";
      }

      for (const col of featureCols) {
        const value = row[col] || "";
        if (!isBlank(value)) attrs[col] = value;
      }

      mainData.set(sku, { attrs, rowNum: row._excel_row });
    }
  }

  const additionalSheets = scoredSheets.filter((s) => s.sheetName !== main.sheetName && s.score > 0);
  const additionalData = new Map<string, { value: string, rowNum: string, sheetName: string, sheetIndex: number }[]>();

  for (const sheet of additionalSheets) {
    if (sheet.headers.length === 0 || !sheet.skuCol) continue;
    
    const activeValueCols = sheet.headers.filter((h) => h !== sheet.skuCol && bulletRegex.test(h));
    
    for (const row of sheet.rows) {
      const sku = row[sheet.skuCol];
      if (isBlank(sku) || !mainData.has(sku)) continue;
      
      if (!additionalData.has(sku)) additionalData.set(sku, []);
      const skuList = additionalData.get(sku)!;

      for (const col of activeValueCols) {
        const val = row[col];
        if (isBlank(val)) continue;
        
        if (!skuList.some(item => item.value === val.trim())) {
          skuList.push({ 
            value: val.trim(), 
            rowNum: row._excel_row, 
            sheetName: sheet.sheetName,
            sheetIndex: sheet.sheetIndex
          });
        }
      }
    }
  }

  const resultRows: TransposedRow[] = [];
  const seenKeys = new Set<string>();

  const getBaseHeading = (h: string) => h.replace(/\s*\d+$/g, "").trim();

  for (const [sku, { attrs }] of mainData.entries()) {
    for (const [heading, value] of Object.entries(attrs)) {
      const compositeKey = `${sku}||${heading}||${value}`;
      const isCorrupted = EXCEL_ERRORS.includes(value);
      const isDuplicate = seenKeys.has(compositeKey);
      seenKeys.add(compositeKey);
      
      resultRows.push({
        sku,
        base_heading: getBaseHeading(heading),
        attribute_heading: heading,
        value,
        validation_status: isCorrupted ? "Corrupted_Value" : isDuplicate ? "Duplicate_Key" : "Valid",
      });
    }

    const additionalValues = additionalData.get(sku);
    if (additionalValues) {
      for (const item of additionalValues) {
        const heading = "Additional Feature Bullet";
        const compositeKey = `${sku}||${heading}||${item.value}`;
        const isCorrupted = EXCEL_ERRORS.includes(item.value);
        const isDuplicate = seenKeys.has(compositeKey);
        seenKeys.add(compositeKey);
        
        resultRows.push({
          sku,
          base_heading: getBaseHeading(heading),
          attribute_heading: heading,
          value: item.value,
          validation_status: isCorrupted ? "Corrupted_Value" : isDuplicate ? "Duplicate_Key" : "Valid",
        });
      }
    }
  }

  return resultRows;
}

export function exportToXlsx(data: TransposedRow[], originalFileName: string): void {
  const exportData = data.map((r) => ({
    SKU: r.sku,
    "Base Heading": r.base_heading,
    "Attribute Heading": r.attribute_heading,
    Value: r.value,
  }));

  const ws = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Transposed");

  const baseName = originalFileName.replace(/\.xlsx$/i, "");
  XLSX.writeFile(wb, `Transposed_${baseName}.xlsx`);
}
