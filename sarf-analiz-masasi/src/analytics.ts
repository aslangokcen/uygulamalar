import { PRODUCT_PRIORITY } from "./data";
import type {
  AnnualTotalRecord,
  ClassifiedConsumptionRecord,
  ConsumptionRecord,
  FocusProductConfig,
  ProductMatch,
  ProductionRecord,
  PurchaseRecord,
} from "./types";
import { formatNumber, normalizeText, safeDivide } from "./utils";

interface KeywordRule {
  product: string;
  analysisUnit: ProductMatch["analysisUnit"];
  keyword: string;
  normalizedKeyword: string;
  priority: number;
}

export interface GroupTotal {
  quantity: number;
  amount: number;
  rows: number;
}

export interface ValidationRow {
  key: string;
  year: number | null;
  factory: string;
  productCode: string;
  productName: string;
  unit: string;
  detailQuantity: number;
  annualQuantity: number;
  diff: number;
  status: "OK" | "Fark" | "Satır yok" | "Yıllık yok";
}

export interface FocusValidationRow {
  key: string;
  year: number | null;
  factory: string;
  product: string;
  unit: string;
  detailQuantity: number;
  annualQuantity: number;
  diff: number;
  status: "OK" | "Fark" | "Satır yok" | "Yıllık yok";
}

export interface PurchaseValidationRow {
  key: string;
  year: number | null;
  productCode: string;
  productName: string;
  unit: string;
  consumptionQuantity: number;
  purchaseQuantity: number;
  diff: number;
  status: "OK" | "Eksik satın alma" | "Fazla satın alma";
}

function productPriority(product: string): number {
  const index = PRODUCT_PRIORITY.indexOf(product);
  return index === -1 ? 100 : index;
}

function buildRules(configs: FocusProductConfig[]): KeywordRule[] {
  return configs
    .flatMap((config) =>
      config.keywords
        .map((keyword) => keyword.trim())
        .filter(Boolean)
        .map((keyword) => ({
          product: config.product,
          analysisUnit: config.analysisUnit,
          keyword,
          normalizedKeyword: normalizeText(keyword),
          priority: productPriority(config.product),
        })),
    )
    .filter((rule) => rule.normalizedKeyword)
    .sort((a, b) => {
      if (a.priority !== b.priority) return a.priority - b.priority;
      if (b.normalizedKeyword.length !== a.normalizedKeyword.length) return b.normalizedKeyword.length - a.normalizedKeyword.length;
      return a.product.localeCompare(b.product, "tr");
    });
}

function containsKeyword(normalizedText: string, normalizedKeyword: string): boolean {
  if (!normalizedKeyword) return false;
  const paddedText = ` ${normalizedText} `;
  const paddedKeyword = ` ${normalizedKeyword} `;
  if (normalizedKeyword.length <= 4 || normalizedKeyword.split(" ").length > 1) {
    return paddedText.includes(paddedKeyword);
  }
  return normalizedText.includes(normalizedKeyword);
}

export function classifyProduct(productName: string, configs: FocusProductConfig[]): ProductMatch | null {
  const normalizedProductName = normalizeText(productName);
  if (!normalizedProductName) return null;
  for (const rule of buildRules(configs)) {
    if (containsKeyword(normalizedProductName, rule.normalizedKeyword)) {
      return {
        product: rule.product,
        analysisUnit: rule.analysisUnit,
        keyword: rule.keyword,
      };
    }
  }
  return null;
}

export function classifyConsumption(records: ConsumptionRecord[], configs: FocusProductConfig[]): ClassifiedConsumptionRecord[] {
  const rules = buildRules(configs);
  return records.map((record) => {
    const normalizedProductName = normalizeText(record.productName);
    const match = rules.find((rule) => containsKeyword(normalizedProductName, rule.normalizedKeyword));
    return {
      ...record,
      focusProduct: match?.product ?? null,
      analysisUnit: match?.analysisUnit ?? null,
      matchedKeyword: match?.keyword ?? null,
    };
  });
}

export function addToMap(map: Map<string, GroupTotal>, key: string, quantity: number, amount: number | null | undefined): void {
  const current = map.get(key) ?? { quantity: 0, amount: 0, rows: 0 };
  current.quantity += Number.isFinite(quantity) ? quantity : 0;
  current.amount += Number.isFinite(amount ?? NaN) ? Number(amount) : 0;
  current.rows += 1;
  map.set(key, current);
}

export function consumptionByProductYear(records: ClassifiedConsumptionRecord[]): Map<string, GroupTotal> {
  const map = new Map<string, GroupTotal>();
  for (const record of records) {
    if (!record.focusProduct || !record.year) continue;
    addToMap(map, `${record.focusProduct}|${record.year}`, record.quantity, record.totalAmount ?? record.amount);
  }
  return map;
}

export function consumptionByProductFactoryYearMonth(records: ClassifiedConsumptionRecord[]): Map<string, GroupTotal> {
  const map = new Map<string, GroupTotal>();
  for (const record of records) {
    if (!record.focusProduct || !record.year || !record.month) continue;
    addToMap(
      map,
      `${record.focusProduct}|${record.factory}|${record.year}|${record.month}`,
      record.quantity,
      record.totalAmount ?? record.amount,
    );
  }
  return map;
}

export function productionByFactoryYearMonth(records: ProductionRecord[]): Map<string, number> {
  const map = new Map<string, number>();
  for (const record of records) {
    if (!record.year) continue;
    const key = `${record.factory}|${record.year}|${record.month}|${record.analysisUnit}`;
    map.set(key, (map.get(key) ?? 0) + record.quantity);
  }
  return map;
}

export function yearsFromData(...lists: Array<Array<{ year: number | null }>>): number[] {
  const years = new Set<number>();
  for (const list of lists) {
    for (const item of list) {
      if (item.year) years.add(item.year);
    }
  }
  return [...years].sort((a, b) => a - b);
}

export function factoriesFromData(records: Array<{ factory: string }>): string[] {
  return [...new Set(records.map((record) => record.factory).filter(Boolean))].sort((a, b) => a.localeCompare(b, "tr"));
}

function exactKey(year: number | null, factory: string, productCode: string, unit: string): string {
  return `${year ?? ""}|${factory}|${normalizeText(productCode)}|${normalizeText(unit)}`;
}

function groupDetailExact(
  records: ConsumptionRecord[],
  scope: "actual" | "source",
): Map<string, { quantity: number; name: string; code: string; year: number | null; factory: string; unit: string }> {
  const map = new Map<string, { quantity: number; name: string; code: string; year: number | null; factory: string; unit: string }>();
  for (const record of records) {
    if (!record.productCode && !record.productName) continue;
    const factory = scope === "source" ? record.sourceFactory : record.factory;
    const code = record.productCode || record.productName;
    const key = exactKey(record.year, factory, code, record.unit);
    const current = map.get(key) ?? {
      quantity: 0,
      name: record.productName,
      code,
      year: record.year,
      factory,
      unit: record.unit,
    };
    current.quantity += record.quantity;
    map.set(key, current);
  }
  return map;
}

function groupAnnualExact(records: AnnualTotalRecord[]): Map<string, { quantity: number; name: string; code: string; year: number | null; factory: string; unit: string }> {
  const map = new Map<string, { quantity: number; name: string; code: string; year: number | null; factory: string; unit: string }>();
  for (const record of records) {
    if (!record.productCode && !record.productName) continue;
    const code = record.productCode || record.productName;
    const key = exactKey(record.year, record.factory, code, record.unit);
    const current = map.get(key) ?? {
      quantity: 0,
      name: record.productName,
      code,
      year: record.year,
      factory: record.factory,
      unit: record.unit,
    };
    current.quantity += record.quantity;
    map.set(key, current);
  }
  return map;
}

export function buildExactValidationRows(consumption: ConsumptionRecord[], annualTotals: AnnualTotalRecord[]): ValidationRow[] {
  const detailActualMap = groupDetailExact(consumption, "actual");
  const detailSourceMap = groupDetailExact(consumption, "source");
  const annualMap = groupAnnualExact(annualTotals);
  const annualFactoryYears = new Set(annualTotals.map((row) => `${row.year ?? ""}|${row.factory}`));
  const rows: ValidationRow[] = [];
  const usedDetailKeys = new Set<string>();

  for (const key of annualMap.keys()) {
    const detail = detailActualMap.get(key) ?? detailSourceMap.get(key);
    const annual = annualMap.get(key);
    const detailQuantity = detail?.quantity ?? 0;
    const annualQuantity = annual?.quantity ?? 0;
    const diff = detailQuantity - annualQuantity;
    const source = annual ?? detail;
    if (!source) continue;
    if (detail) usedDetailKeys.add(exactKey(detail.year, detail.factory, detail.code, detail.unit));
    rows.push({
      key,
      year: source.year,
      factory: source.factory,
      productCode: source.code,
      productName: source.name,
      unit: source.unit,
      detailQuantity,
      annualQuantity,
      diff,
      status: Math.abs(diff) <= 0.01 ? "OK" : detail ? (annual ? "Fark" : "Yıllık yok") : "Satır yok",
    });
  }

  for (const [key, detail] of detailActualMap.entries()) {
    if (usedDetailKeys.has(key) || annualMap.has(key)) continue;
    if (!annualFactoryYears.has(`${detail.year ?? ""}|${detail.factory}`)) continue;
    rows.push({
      key,
      year: detail.year,
      factory: detail.factory,
      productCode: detail.code,
      productName: detail.name,
      unit: detail.unit,
      detailQuantity: detail.quantity,
      annualQuantity: 0,
      diff: detail.quantity,
      status: "Yıllık yok",
    });
  }

  return rows.sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));
}

export function buildFocusValidationRows(
  consumption: ClassifiedConsumptionRecord[],
  annualTotals: AnnualTotalRecord[],
  configs: FocusProductConfig[],
): FocusValidationRow[] {
  const detailMap = new Map<string, { quantity: number; unit: string; year: number | null; factory: string; product: string }>();
  for (const record of consumption) {
    if (!record.focusProduct || !record.year) continue;
    const key = `${record.year}|${record.factory}|${record.focusProduct}|${record.unit}`;
    const current = detailMap.get(key) ?? {
      quantity: 0,
      unit: record.unit,
      year: record.year,
      factory: record.factory,
      product: record.focusProduct,
    };
    current.quantity += record.quantity;
    detailMap.set(key, current);
  }

  const annualMap = new Map<string, { quantity: number; unit: string; year: number | null; factory: string; product: string }>();
  const annualFactoryYears = new Set<string>();
  for (const record of annualTotals) {
    const match = classifyProduct(record.productName, configs);
    if (!match || !record.year) continue;
    annualFactoryYears.add(`${record.year}|${record.factory}`);
    const key = `${record.year}|${record.factory}|${match.product}|${record.unit}`;
    const current = annualMap.get(key) ?? {
      quantity: 0,
      unit: record.unit,
      year: record.year,
      factory: record.factory,
      product: match.product,
    };
    current.quantity += record.quantity;
    annualMap.set(key, current);
  }

  const keys = new Set([...annualMap.keys(), ...[...detailMap.keys()].filter((key) => {
    const [year, factory] = key.split("|");
    return annualFactoryYears.has(`${year}|${factory}`);
  })]);
  return [...keys]
    .map((key) => {
      const detail = detailMap.get(key);
      const annual = annualMap.get(key);
      const source = annual ?? detail!;
      const diff = (detail?.quantity ?? 0) - (annual?.quantity ?? 0);
      return {
        key,
        year: source.year,
        factory: source.factory,
        product: source.product,
        unit: source.unit,
        detailQuantity: detail?.quantity ?? 0,
        annualQuantity: annual?.quantity ?? 0,
        diff,
        status: Math.abs(diff) <= 0.01 ? "OK" : detail ? (annual ? "Fark" : "Yıllık yok") : "Satır yok",
      } satisfies FocusValidationRow;
    })
    .sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));
}

export function buildPurchaseValidationRows(consumption: ConsumptionRecord[], purchases: PurchaseRecord[]): PurchaseValidationRow[] {
  const consumptionMap = new Map<string, { quantity: number; name: string; code: string; year: number | null; unit: string }>();
  for (const record of consumption) {
    if (!record.productCode && !record.productName) continue;
    const code = record.productCode || record.productName;
    const key = `${record.year ?? ""}|${normalizeText(code)}|${normalizeText(record.unit)}`;
    const current = consumptionMap.get(key) ?? { quantity: 0, name: record.productName, code, year: record.year, unit: record.unit };
    current.quantity += record.quantity;
    consumptionMap.set(key, current);
  }

  const purchaseMap = new Map<string, { quantity: number; name: string; code: string; year: number | null; unit: string }>();
  for (const record of purchases) {
    if (!record.productCode && !record.productName) continue;
    const code = record.productCode || record.productName;
    const key = `${record.year ?? ""}|${normalizeText(code)}|${normalizeText(record.unit)}`;
    const current = purchaseMap.get(key) ?? { quantity: 0, name: record.productName, code, year: record.year, unit: record.unit };
    current.quantity += record.quantity;
    purchaseMap.set(key, current);
  }

  return [...new Set([...consumptionMap.keys(), ...purchaseMap.keys()])]
    .map((key) => {
      const consumptionItem = consumptionMap.get(key);
      const purchaseItem = purchaseMap.get(key);
      const source = purchaseItem ?? consumptionItem!;
      const consumptionQuantity = consumptionItem?.quantity ?? 0;
      const purchaseQuantity = purchaseItem?.quantity ?? 0;
      const diff = purchaseQuantity - consumptionQuantity;
      return {
        key,
        year: source.year,
        productCode: source.code,
        productName: source.name,
        unit: source.unit,
        consumptionQuantity,
        purchaseQuantity,
        diff,
        status: Math.abs(diff) <= 0.01 ? "OK" : diff < 0 ? "Eksik satın alma" : "Fazla satın alma",
      } satisfies PurchaseValidationRow;
    })
    .sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));
}

export function metricForCell(options: {
  consumptionMap: Map<string, GroupTotal>;
  productionMap: Map<string, number>;
  product: string;
  factory: string;
  year: number;
  month: number | null;
  analysisUnit: ProductMatch["analysisUnit"];
  metric: "quantity" | "amount" | "production" | "ratio";
}): number | null {
  const months = options.month ? [options.month] : Array.from({ length: 12 }, (_, index) => index + 1);
  let consumptionQuantity = 0;
  let consumptionAmount = 0;
  let productionQuantity = 0;

  for (const month of months) {
    const total = options.consumptionMap.get(`${options.product}|${options.factory}|${options.year}|${month}`);
    consumptionQuantity += total?.quantity ?? 0;
    consumptionAmount += total?.amount ?? 0;
    productionQuantity += options.productionMap.get(`${options.factory}|${options.year}|${month}|${options.analysisUnit}`) ?? 0;
  }

  if (options.metric === "quantity") return consumptionQuantity;
  if (options.metric === "amount") return consumptionAmount;
  if (options.metric === "production") return productionQuantity;
  return safeDivide(consumptionQuantity, productionQuantity);
}

export function exportRowsForConsumption(records: ClassifiedConsumptionRecord[]): Array<Record<string, unknown>> {
  return records.map((record) => ({
    Yıl: record.year ?? "",
    Ay: record.month ?? "",
    Gün: record.day ?? "",
    Fabrika: record.factory,
    Departman: record.department,
    Ambar: record.warehouse,
    Tedarikçi: record.supplier,
    ÜrünKodu: record.productCode,
    ÜrünAdı: record.productName,
    OdakÜrün: record.focusProduct ?? "",
    EşleşenKelime: record.matchedKeyword ?? "",
    Miktar: formatNumber(record.quantity, 4),
    Birim: record.unit,
    BirimFiyat: record.unitPrice ?? "",
    ToplamTutar: record.totalAmount ?? record.amount ?? "",
    Dosya: record.fileName,
  }));
}
