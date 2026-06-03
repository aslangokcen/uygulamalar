import * as XLSX from "xlsx";
import { MONTHS } from "./data";
import type {
  AnnualTotalRecord,
  ConsumptionRecord,
  FileKind,
  ParsedBatch,
  ParsedFileSummary,
  ProductionRecord,
  PurchaseRecord,
} from "./types";
import {
  cleanText,
  extractYear,
  filePath,
  inferFactory,
  makeFileKey,
  monthValueFromHeader,
  normalizeFactory,
  normalizeText,
  normalizeUnit,
  productionUnitFromLabel,
  toNumber,
} from "./utils";

type Row = Record<string, unknown>;

interface SheetContext {
  file: File;
  fileKey: string;
  path: string;
  sheetName: string;
  kind: FileKind;
  headers: string[];
  rows: Row[];
  warnings: string[];
}

function pushMany<T>(target: T[], source: T[]): void {
  for (const item of source) target.push(item);
}

const HEADER_CANDIDATES = {
  date: ["İrsaliye Tarihi", "Tarih", "Fiş Tarihi", "Belge Tarihi", "Fatura Tarihi", "Sipariş Tarihi"],
  documentNo: ["Fiş No", "Fis No", "Belge No", "İrsaliye No", "Irsaliye No", "Fatura No", "Sipariş No"],
  documentType: ["Fiş Türü", "Fis Turu", "Belge Türü", "Hareket Türü"],
  supplier: ["Cari Hesap", "Cari Adı", "Cari Unvan", "Tedarikçi", "Tedarikci", "Satıcı", "Satici"],
  factory: ["Fabrika", "Tesis", "Lokasyon", "İşyeri", "Isyeri"],
  department: ["Departman", "Bölüm", "Bolum", "Masraf Merkezi"],
  warehouse: ["Ambar", "Depo"],
  productCode: ["Ürün Kodu", "Urun Kodu", "Malzeme Kodu", "Stok Kodu", "Kod"],
  productName: ["Ürün Adı", "Urun Adi", "Malzeme Adı", "Malzeme Adi", "Stok Adı", "Stok Adi", "Açıklama"],
  quantity: ["Miktar", "Sarf Miktarı", "Sarf Miktari", "Toplam Miktar", "Giren Miktar"],
  unit: ["Birim", "Ölçü Birimi", "Olcu Birimi"],
  unitPrice: ["Birim Fiyat", "Fiyat", "Birim Maliyet"],
  amount: ["Tutar", "Net Tutar", "Satır Tutarı", "Satir Tutari"],
  totalAmount: ["Toplam Tutar", "Genel Tutar", "Toplam"],
};

function normalizedHeaders(headers: string[]): Array<[string, string]> {
  return headers.map((header) => [header, normalizeText(header)]);
}

function findHeader(headers: string[], candidates: string[]): string | null {
  const prepared = normalizedHeaders(headers);
  const targets = candidates.map((candidate) => normalizeText(candidate));
  for (const target of targets) {
    const exact = prepared.find(([, normalized]) => normalized === target);
    if (exact) return exact[0];
  }
  for (const target of targets) {
    const partial = prepared.find(([, normalized]) => normalized.includes(target) || target.includes(normalized));
    if (partial) return partial[0];
  }
  return null;
}

function getValue(row: Row, header: string | null): unknown {
  return header ? row[header] : null;
}

function getText(row: Row, header: string | null): string {
  return cleanText(getValue(row, header));
}

function headersFromRows(rows: Row[]): string[] {
  const seen = new Set<string>();
  for (const row of rows.slice(0, 10)) {
    Object.keys(row).forEach((header) => seen.add(header));
  }
  return [...seen];
}

function detectKind(headers: string[], sourceName: string): FileKind {
  const monthColumns = headers.filter((header) => monthValueFromHeader(header)).length;
  const hasFactory = Boolean(findHeader(headers, HEADER_CANDIDATES.factory));
  const hasUnit = Boolean(findHeader(headers, HEADER_CANDIDATES.unit));
  const hasDate = Boolean(findHeader(headers, HEADER_CANDIDATES.date));
  const hasProduct = Boolean(findHeader(headers, HEADER_CANDIDATES.productName));
  const hasProductCode = Boolean(findHeader(headers, HEADER_CANDIDATES.productCode));
  const hasQuantity = Boolean(findHeader(headers, HEADER_CANDIDATES.quantity));
  const source = normalizeText(sourceName);

  if (hasFactory && hasUnit && monthColumns >= 6) return "production";
  if (hasDate && hasProduct && hasQuantity) return "consumption";
  if (source.includes("SATIN") || source.includes("ALMA") || source.includes("SIPARIS")) return "purchase";
  if (hasProduct && hasProductCode && hasQuantity && hasUnit && !hasDate) return "annualTotal";
  if (hasProduct && hasQuantity && hasUnit) return "annualTotal";
  return "unknown";
}

function asDate(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  const text = cleanText(value);
  if (!text) return null;

  const iso = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));

  const local = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})/);
  if (local) {
    const rawYear = Number(local[3]);
    const year = rawYear < 100 ? 2000 + rawYear : rawYear;
    return new Date(year, Number(local[2]) - 1, Number(local[1]));
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function rowHasData(row: Row): boolean {
  return Object.values(row).some((value) => value !== null && value !== undefined && cleanText(value) !== "");
}

function parseConsumptionSheet(context: SheetContext): ConsumptionRecord[] {
  const { headers, rows, path, file, fileKey, sheetName } = context;
  const dateHeader = findHeader(headers, HEADER_CANDIDATES.date);
  const documentNoHeader = findHeader(headers, HEADER_CANDIDATES.documentNo);
  const documentTypeHeader = findHeader(headers, HEADER_CANDIDATES.documentType);
  const supplierHeader = findHeader(headers, HEADER_CANDIDATES.supplier);
  const factoryHeader = findHeader(headers, HEADER_CANDIDATES.factory);
  const departmentHeader = findHeader(headers, HEADER_CANDIDATES.department);
  const warehouseHeader = findHeader(headers, HEADER_CANDIDATES.warehouse);
  const productCodeHeader = findHeader(headers, HEADER_CANDIDATES.productCode);
  const productNameHeader = findHeader(headers, HEADER_CANDIDATES.productName);
  const quantityHeader = findHeader(headers, HEADER_CANDIDATES.quantity);
  const unitHeader = findHeader(headers, HEADER_CANDIDATES.unit);
  const unitPriceHeader = findHeader(headers, HEADER_CANDIDATES.unitPrice);
  const amountHeader = findHeader(headers, HEADER_CANDIDATES.amount);
  const totalAmountHeader = findHeader(headers, HEADER_CANDIDATES.totalAmount);
  const fallbackYear = extractYear(path);
  const sourceFactory = inferFactory(path);

  return rows.flatMap((row, index) => {
    const productName = getText(row, productNameHeader);
    const quantity = toNumber(getValue(row, quantityHeader));
    if (!productName || quantity === null) return [];

    const date = asDate(getValue(row, dateHeader));
    const year = date?.getFullYear() ?? fallbackYear;
    const month = date ? date.getMonth() + 1 : null;
    const day = date ? date.getDate() : null;
    const rowFactory = normalizeFactory(getValue(row, factoryHeader), sourceFactory || "Bilinmiyor");

    return [
      {
        id: `${fileKey}:${sheetName}:C${index + 2}`,
        fileKey,
        fileName: file.name,
        sourceFactory: sourceFactory || rowFactory,
        sheetName,
        rowNumber: index + 2,
        date,
        year,
        month,
        day,
        factory: rowFactory,
        department: getText(row, departmentHeader),
        warehouse: getText(row, warehouseHeader),
        supplier: getText(row, supplierHeader),
        documentNo: getText(row, documentNoHeader),
        documentType: getText(row, documentTypeHeader),
        productCode: getText(row, productCodeHeader),
        productName,
        quantity,
        unit: normalizeUnit(getValue(row, unitHeader)),
        unitPrice: toNumber(getValue(row, unitPriceHeader)),
        amount: toNumber(getValue(row, amountHeader)),
        totalAmount: toNumber(getValue(row, totalAmountHeader)),
      },
    ];
  });
}

function parseProductionSheet(context: SheetContext): ProductionRecord[] {
  const { headers, rows, path, file, fileKey, sheetName } = context;
  const factoryHeader = findHeader(headers, HEADER_CANDIDATES.factory);
  const unitHeader = findHeader(headers, HEADER_CANDIDATES.unit);
  const year = extractYear(path);
  const sourceFactory = inferFactory(path);
  const monthHeaders = headers
    .map((header) => ({ header, month: monthValueFromHeader(header) }))
    .filter((item): item is { header: string; month: number } => item.month !== null);

  return rows.flatMap((row, index) => {
    const sourceLabel = getText(row, unitHeader);
    const analysisUnit = productionUnitFromLabel(sourceLabel);
    if (!analysisUnit) return [];
    const factory = normalizeFactory(getValue(row, factoryHeader), sourceFactory || "Bilinmiyor");

    return monthHeaders.flatMap(({ header, month }) => {
      const quantity = toNumber(row[header]);
      if (quantity === null) return [];
      return [
        {
          id: `${fileKey}:${sheetName}:P${index + 2}:${month}`,
          fileKey,
          fileName: file.name,
          sourceFactory: sourceFactory || factory,
          sheetName,
          rowNumber: index + 2,
          year,
          month,
          factory,
          analysisUnit,
          sourceLabel,
          quantity,
        },
      ];
    });
  });
}

function parseAnnualTotalSheet(context: SheetContext): AnnualTotalRecord[] {
  const { headers, rows, path, file, fileKey, sheetName } = context;
  const factoryHeader = findHeader(headers, HEADER_CANDIDATES.factory);
  const productCodeHeader = findHeader(headers, HEADER_CANDIDATES.productCode);
  const productNameHeader = findHeader(headers, HEADER_CANDIDATES.productName);
  const quantityHeader = findHeader(headers, HEADER_CANDIDATES.quantity);
  const unitHeader = findHeader(headers, HEADER_CANDIDATES.unit);
  const amountHeader = findHeader(headers, HEADER_CANDIDATES.amount);
  const totalAmountHeader = findHeader(headers, HEADER_CANDIDATES.totalAmount);
  const year = extractYear(path);
  const sourceFactory = inferFactory(path);

  return rows.flatMap((row, index) => {
    const productName = getText(row, productNameHeader);
    const quantity = toNumber(getValue(row, quantityHeader));
    if (!productName || quantity === null) return [];
    const factory = normalizeFactory(getValue(row, factoryHeader), sourceFactory || "Bilinmiyor");

    return [
      {
        id: `${fileKey}:${sheetName}:A${index + 2}`,
        fileKey,
        fileName: file.name,
        sourceFactory: sourceFactory || factory,
        sheetName,
        rowNumber: index + 2,
        year,
        factory,
        productCode: getText(row, productCodeHeader),
        productName,
        quantity,
        unit: normalizeUnit(getValue(row, unitHeader)),
        amount: toNumber(getValue(row, amountHeader)),
        totalAmount: toNumber(getValue(row, totalAmountHeader)),
      },
    ];
  });
}

function parsePurchaseSheet(context: SheetContext): PurchaseRecord[] {
  const { headers, rows, path, file, fileKey, sheetName } = context;
  const dateHeader = findHeader(headers, HEADER_CANDIDATES.date);
  const supplierHeader = findHeader(headers, HEADER_CANDIDATES.supplier);
  const factoryHeader = findHeader(headers, HEADER_CANDIDATES.factory);
  const productCodeHeader = findHeader(headers, HEADER_CANDIDATES.productCode);
  const productNameHeader = findHeader(headers, HEADER_CANDIDATES.productName);
  const quantityHeader = findHeader(headers, HEADER_CANDIDATES.quantity);
  const unitHeader = findHeader(headers, HEADER_CANDIDATES.unit);
  const unitPriceHeader = findHeader(headers, HEADER_CANDIDATES.unitPrice);
  const totalAmountHeader = findHeader(headers, HEADER_CANDIDATES.totalAmount) ?? findHeader(headers, HEADER_CANDIDATES.amount);
  const fallbackYear = extractYear(path);
  const sourceFactory = inferFactory(path);

  return rows.flatMap((row, index) => {
    const productName = getText(row, productNameHeader);
    const quantity = toNumber(getValue(row, quantityHeader));
    if (!productName || quantity === null) return [];
    const date = asDate(getValue(row, dateHeader));
    const factory = normalizeFactory(getValue(row, factoryHeader), sourceFactory || "Bilinmiyor");

    return [
      {
        id: `${fileKey}:${sheetName}:S${index + 2}`,
        fileKey,
        fileName: file.name,
        sourceFactory: sourceFactory || factory,
        sheetName,
        rowNumber: index + 2,
        date,
        year: date?.getFullYear() ?? fallbackYear,
        month: date ? date.getMonth() + 1 : null,
        factory,
        supplier: getText(row, supplierHeader),
        productCode: getText(row, productCodeHeader),
        productName,
        quantity,
        unit: normalizeUnit(getValue(row, unitHeader)),
        unitPrice: toNumber(getValue(row, unitPriceHeader)),
        totalAmount: toNumber(getValue(row, totalAmountHeader)),
      },
    ];
  });
}

function summarizeKind(kind: FileKind): string {
  switch (kind) {
    case "consumption":
      return "satır bazlı sarf";
    case "production":
      return "üretim";
    case "annualTotal":
      return "ürün bazında yıllık";
    case "purchase":
      return "satın alma";
    default:
      return "tanınmadı";
  }
}

export async function parseWorkbookFile(file: File): Promise<ParsedBatch> {
  const fileKey = makeFileKey(file);
  const path = filePath(file);
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", cellDates: true });
  const batch: ParsedBatch = {
    files: [],
    consumption: [],
    production: [],
    annualTotals: [],
    purchases: [],
  };

  const workbookWarnings: string[] = [];
  let workbookKind: FileKind = "unknown";
  let totalRows = 0;
  const usedSheets: string[] = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json<Row>(sheet, { defval: null, raw: true, blankrows: false }).filter(rowHasData);
    if (!rows.length) continue;

    const headers = headersFromRows(rows);
    const kind = detectKind(headers, `${path} ${sheetName}`);
    if (workbookKind === "unknown") workbookKind = kind;
    if (kind !== workbookKind && kind !== "unknown") {
      workbookWarnings.push(`${sheetName}: ${summarizeKind(kind)} olarak algılandı`);
    }

    const context: SheetContext = {
      file,
      fileKey,
      path,
      sheetName,
      kind,
      headers,
      rows,
      warnings: workbookWarnings,
    };

    if (kind === "consumption") pushMany(batch.consumption, parseConsumptionSheet(context));
    if (kind === "production") pushMany(batch.production, parseProductionSheet(context));
    if (kind === "annualTotal") pushMany(batch.annualTotals, parseAnnualTotalSheet(context));
    if (kind === "purchase") pushMany(batch.purchases, parsePurchaseSheet(context));
    if (kind === "unknown") workbookWarnings.push(`${sheetName}: kolon şablonu tanınmadı`);

    totalRows += rows.length;
    usedSheets.push(sheetName);
  }

  if (!usedSheets.length) workbookWarnings.push("Okunabilir satır bulunamadı");
  if (workbookKind === "unknown" && inferFactory(path) && extractYear(path)) {
    workbookWarnings.push("Dosya adı yıl/fabrika içeriyor ama kolonlar eşleşmedi");
  }

  const summary: ParsedFileSummary = {
    key: fileKey,
    name: file.name,
    path,
    size: file.size,
    kind: workbookKind,
    sheets: usedSheets,
    rows:
      workbookKind === "consumption"
        ? batch.consumption.length
        : workbookKind === "production"
          ? batch.production.length
          : workbookKind === "annualTotal"
            ? batch.annualTotals.length
            : workbookKind === "purchase"
              ? batch.purchases.length
              : totalRows,
    warnings: workbookWarnings,
  };

  batch.files.push(summary);
  return batch;
}

export async function parseWorkbookFiles(files: File[], onProgress?: (message: string) => void): Promise<ParsedBatch> {
  const combined: ParsedBatch = {
    files: [],
    consumption: [],
    production: [],
    annualTotals: [],
    purchases: [],
  };

  for (let index = 0; index < files.length; index += 1) {
    const file = files[index];
    onProgress?.(`${index + 1}/${files.length} okunuyor: ${file.name}`);
    const parsed = await parseWorkbookFile(file);
    pushMany(combined.files, parsed.files);
    pushMany(combined.consumption, parsed.consumption);
    pushMany(combined.production, parsed.production);
    pushMany(combined.annualTotals, parsed.annualTotals);
    pushMany(combined.purchases, parsed.purchases);
    await new Promise((resolve) => window.setTimeout(resolve, 0));
  }

  onProgress?.(`${files.length} dosya işlendi`);
  return combined;
}

export function expectedProductionColumns(): string {
  return MONTHS.map((month) => month.short).join(", ");
}
