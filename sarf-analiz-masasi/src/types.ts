export type FileKind = "consumption" | "production" | "annualTotal" | "purchase" | "unknown";

export type AnalysisUnit = "M2" | "KG" | "Palet Sayısı";

export interface FocusProductConfig {
  product: string;
  analysisUnit: AnalysisUnit;
  keywords: string[];
}

export interface ParsedFileSummary {
  key: string;
  name: string;
  path: string;
  size: number;
  kind: FileKind;
  sheets: string[];
  rows: number;
  warnings: string[];
}

export interface ConsumptionRecord {
  id: string;
  fileKey: string;
  fileName: string;
  sourceFactory: string;
  sheetName: string;
  rowNumber: number;
  date: Date | null;
  year: number | null;
  month: number | null;
  day: number | null;
  factory: string;
  department: string;
  warehouse: string;
  supplier: string;
  documentNo: string;
  documentType: string;
  productCode: string;
  productName: string;
  quantity: number;
  unit: string;
  unitPrice: number | null;
  amount: number | null;
  totalAmount: number | null;
}

export interface ClassifiedConsumptionRecord extends ConsumptionRecord {
  focusProduct: string | null;
  analysisUnit: AnalysisUnit | null;
  matchedKeyword: string | null;
}

export interface ProductionRecord {
  id: string;
  fileKey: string;
  fileName: string;
  sourceFactory: string;
  sheetName: string;
  rowNumber: number;
  year: number | null;
  month: number;
  factory: string;
  analysisUnit: AnalysisUnit;
  sourceLabel: string;
  quantity: number;
}

export interface AnnualTotalRecord {
  id: string;
  fileKey: string;
  fileName: string;
  sourceFactory: string;
  sheetName: string;
  rowNumber: number;
  year: number | null;
  factory: string;
  productCode: string;
  productName: string;
  quantity: number;
  unit: string;
  amount: number | null;
  totalAmount: number | null;
}

export interface PurchaseRecord {
  id: string;
  fileKey: string;
  fileName: string;
  sourceFactory: string;
  sheetName: string;
  rowNumber: number;
  date: Date | null;
  year: number | null;
  month: number | null;
  factory: string;
  supplier: string;
  productCode: string;
  productName: string;
  quantity: number;
  unit: string;
  unitPrice: number | null;
  totalAmount: number | null;
}

export interface ParsedBatch {
  files: ParsedFileSummary[];
  consumption: ConsumptionRecord[];
  production: ProductionRecord[];
  annualTotals: AnnualTotalRecord[];
  purchases: PurchaseRecord[];
}

export interface ProductMatch {
  product: string;
  analysisUnit: AnalysisUnit;
  keyword: string;
}
