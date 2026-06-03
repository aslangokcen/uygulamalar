import { MONTHS } from "./data";
import type { AnalysisUnit } from "./types";

// Tesis/lokasyon takma adlari — kendi tesis adlarinizla degistirin.
// Ornek: girdideki ham metni ("FABRIKA-A" vb.) okunabilir bir etikete eslestirir.
const FACTORY_ALIASES = [
  { test: "TESIS1", label: "Tesis 1" },
  { test: "TESIS2", label: "Tesis 2" },
  { test: "TESIS3", label: "Tesis 3" },
  { test: "TESIS4", label: "Tesis 4" },
  { test: "TESIS5", label: "Tesis 5" },
];

export function normalizeText(value: unknown): string {
  return String(value ?? "")
    .toLocaleUpperCase("tr-TR")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/İ/g, "I")
    .replace(/ı/g, "I")
    .replace(/[^A-Z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

export function cleanText(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

export function makeFileKey(file: File): string {
  return `${file.name}|${file.size}|${file.lastModified}`;
}

export function filePath(file: File): string {
  const withPath = file as File & { webkitRelativePath?: string };
  return withPath.webkitRelativePath || file.name;
}

export function extractYear(text: string): number | null {
  const normalized = normalizeText(text);
  const longYear = normalized.match(/\b(20[0-9]{2})\b/);
  if (longYear) return Number(longYear[1]);

  const shortYear = normalized.match(/(?:^|[^0-9])([2-4][0-9])(?:[^0-9]|$)/);
  if (!shortYear) return null;

  const value = Number(shortYear[1]);
  return value >= 20 && value <= 49 ? 2000 + value : null;
}

export function inferFactory(text: string): string {
  const normalized = normalizeText(text);
  for (const item of FACTORY_ALIASES) {
    if (normalized.includes(item.test)) return item.label;
  }
  return "";
}

export function normalizeFactory(value: unknown, fallback = ""): string {
  const raw = cleanText(value);
  if (!raw) return fallback;
  const normalized = normalizeText(raw.replace(/Fabrika/gi, ""));
  for (const item of FACTORY_ALIASES) {
    if (normalized.includes(item.test)) return item.label;
  }
  return raw.replace(/\s*Fabrika\s*/gi, "").trim() || fallback;
}

export function toNumber(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value === "boolean" || value === null || value === undefined) return null;
  let text = String(value).trim();
  if (!text) return null;
  text = text.replace(/\s/g, "");
  if (text.includes(",") && text.includes(".")) {
    text = text.replace(/\./g, "").replace(",", ".");
  } else if (text.includes(",")) {
    text = text.replace(",", ".");
  }
  text = text.replace(/[^0-9.-]/g, "");
  const number = Number(text);
  return Number.isFinite(number) ? number : null;
}

export function normalizeUnit(value: unknown): string {
  const unit = normalizeText(value);
  if (unit === "M2" || unit === "M 2") return "M2";
  if (unit === "KG" || unit === "KILOGRAM") return "KG";
  if (unit === "ADET") return "ADET";
  if (unit === "LITRE" || unit === "LT") return "LITRE";
  return cleanText(value).toLocaleUpperCase("tr-TR");
}

export function productionUnitFromLabel(value: unknown): AnalysisUnit | null {
  const label = normalizeText(value);
  if (label === "M2") return "M2";
  if (label === "KG") return "KG";
  if (label.includes("PALET")) return "Palet Sayısı";
  return null;
}

export function monthValueFromHeader(header: string): number | null {
  const normalized = normalizeText(header);
  const found = MONTHS.find((month) => normalizeText(month.short) === normalized || normalizeText(month.label) === normalized);
  return found?.value ?? null;
}

export function monthLabel(value: number | null | undefined): string {
  if (!value) return "-";
  return MONTHS.find((month) => month.value === value)?.short ?? String(value);
}

export function formatNumber(value: number | null | undefined, maximumFractionDigits = 2): string {
  if (value === null || value === undefined || !Number.isFinite(value)) return "-";
  return new Intl.NumberFormat("tr-TR", { maximumFractionDigits }).format(value);
}

export function formatRatio(value: number | null | undefined): string {
  if (value === null || value === undefined || !Number.isFinite(value)) return "-";
  return new Intl.NumberFormat("tr-TR", { maximumFractionDigits: 4 }).format(value);
}

export function safeDivide(numerator: number, denominator: number): number | null {
  if (!Number.isFinite(numerator) || !Number.isFinite(denominator) || denominator === 0) return null;
  return numerator / denominator;
}

export function csvEscape(value: unknown): string {
  const text = String(value ?? "");
  if (/[",\n;]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

export function downloadCsv(name: string, rows: Array<Record<string, unknown>>): void {
  if (!rows.length) return;
  const headers = Object.keys(rows[0]);
  const body = [
    headers.join(";"),
    ...rows.map((row) => headers.map((header) => csvEscape(row[header])).join(";")),
  ].join("\n");
  const blob = new Blob(["\ufeff", body], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

export function sum(values: number[]): number {
  return values.reduce((total, value) => total + (Number.isFinite(value) ? value : 0), 0);
}
