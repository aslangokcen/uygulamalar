import { useEffect, useMemo, useRef, useState } from "react";
import {
  AlertTriangle,
  BarChart3,
  CheckCircle2,
  Database,
  Download,
  FileSpreadsheet,
  Filter,
  FolderInput,
  GitCompare,
  Layers3,
  RefreshCw,
  Search,
  Settings2,
  ShieldCheck,
  Trash2,
  Upload,
} from "lucide-react";
import {
  buildExactValidationRows,
  buildFocusValidationRows,
  buildPurchaseValidationRows,
  classifyConsumption,
  consumptionByProductFactoryYearMonth,
  consumptionByProductYear,
  exportRowsForConsumption,
  factoriesFromData,
  metricForCell,
  productionByFactoryYearMonth,
  yearsFromData,
} from "./analytics";
import { ANALYSIS_UNITS, DEFAULT_PRODUCTS, MONTHS } from "./data";
import { parseWorkbookFiles } from "./parser";
import type {
  AnnualTotalRecord,
  ClassifiedConsumptionRecord,
  ConsumptionRecord,
  FileKind,
  FocusProductConfig,
  ParsedFileSummary,
  ProductionRecord,
  PurchaseRecord,
} from "./types";
import { downloadCsv, formatNumber, formatRatio, makeFileKey, monthLabel, safeDivide } from "./utils";

type Tab = "overview" | "compare" | "validation" | "data";
type Metric = "quantity" | "amount" | "production" | "ratio";

const STORAGE_KEY = "sarf-analiz-masasi.products.v1";

const KIND_LABELS: Record<FileKind, string> = {
  consumption: "Satır bazlı sarf",
  production: "Üretim",
  annualTotal: "Ürün bazında yıllık",
  purchase: "Satın alma",
  unknown: "Tanınmadı",
};

function loadProducts(): FocusProductConfig[] {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (!stored) return DEFAULT_PRODUCTS;
    const parsed = JSON.parse(stored) as FocusProductConfig[];
    return DEFAULT_PRODUCTS.map((item) => parsed.find((storedItem) => storedItem.product === item.product) ?? item);
  } catch {
    return DEFAULT_PRODUCTS;
  }
}

function fileSizeLabel(size: number): string {
  if (size > 1024 * 1024) return `${formatNumber(size / (1024 * 1024), 1)} MB`;
  return `${formatNumber(size / 1024, 1)} KB`;
}

function metricLabel(metric: Metric): string {
  if (metric === "quantity") return "Sarf";
  if (metric === "amount") return "Tutar";
  if (metric === "production") return "Üretim";
  return "Sarf / Üretim";
}

function displayMetric(value: number | null, metric: Metric): string {
  if (metric === "ratio") return formatRatio(value);
  return formatNumber(value, metric === "quantity" ? 2 : 0);
}

function statusClass(status: string): string {
  if (status === "OK") return "status ok";
  if (status.includes("yok") || status.includes("Eksik")) return "status warn";
  return "status bad";
}

function KpiCard({
  label,
  value,
  detail,
  tone = "default",
}: {
  label: string;
  value: string;
  detail: string;
  tone?: "default" | "good" | "warn";
}) {
  return (
    <section className={`kpi ${tone}`}>
      <span>{label}</span>
      <strong>{value}</strong>
      <small>{detail}</small>
    </section>
  );
}

function EmptyState({ title, detail }: { title: string; detail: string }) {
  return (
    <div className="empty-state">
      <Database size={30} />
      <strong>{title}</strong>
      <span>{detail}</span>
    </div>
  );
}

function UploadPanel({
  onFiles,
  isLoading,
  progress,
}: {
  onFiles: (files: FileList | null) => void;
  isLoading: boolean;
  progress: string;
}) {
  const inputRef = useRef<HTMLInputElement | null>(null);

  return (
    <div
      className={`upload-panel ${isLoading ? "loading" : ""}`}
      onDragOver={(event) => event.preventDefault()}
      onDrop={(event) => {
        event.preventDefault();
        onFiles(event.dataTransfer.files);
      }}
    >
      <div className="upload-copy">
        <FileSpreadsheet size={24} />
        <div>
          <strong>Excel dosyaları</strong>
          <span>{isLoading ? progress : "Sarf, üretim, yıllık toplam ve satın alma dosyaları"}</span>
        </div>
      </div>
      <button type="button" className="button primary" onClick={() => inputRef.current?.click()} disabled={isLoading}>
        <Upload size={17} />
        Dosya ekle
      </button>
      <input
        ref={inputRef}
        className="hidden-input"
        type="file"
        multiple
        accept=".xlsx,.xls,.csv,.tsv"
        onChange={(event) => {
          onFiles(event.target.files);
          event.currentTarget.value = "";
        }}
      />
    </div>
  );
}

function OverviewPage({
  products,
  records,
  production,
  annualTotals,
  files,
}: {
  products: FocusProductConfig[];
  records: ClassifiedConsumptionRecord[];
  production: ProductionRecord[];
  annualTotals: AnnualTotalRecord[];
  files: ParsedFileSummary[];
}) {
  const matched = records.filter((record) => record.focusProduct);
  const productYearMap = useMemo(() => consumptionByProductYear(records), [records]);
  const exactValidation = useMemo(() => buildExactValidationRows(records, annualTotals), [records, annualTotals]);
  const years = yearsFromData(records, production, annualTotals);
  const totalRows = records.length;
  const matchedRatio = safeDivide(matched.length, totalRows);
  const factories = factoriesFromData(records);
  const productRows = products.map((product) => {
    const yearly = years.map((year) => productYearMap.get(`${product.product}|${year}`)?.quantity ?? 0);
    const total = yearly.reduce((sum, value) => sum + value, 0);
    const rows = records.filter((record) => record.focusProduct === product.product).length;
    return { product, yearly, total, rows };
  });
  const maxProductTotal = Math.max(...productRows.map((row) => row.total), 1);
  const validationProblems = exactValidation.filter((row) => row.status !== "OK").length;

  if (!files.length) {
    return <EmptyState title="Veri bekleniyor" detail="Excel dosyalarını eklediğinde tablo ve kıyaslar burada oluşur." />;
  }

  return (
    <div className="page-grid">
      <div className="kpi-grid">
        <KpiCard label="Dosya" value={formatNumber(files.length, 0)} detail={`${files.filter((file) => file.kind !== "unknown").length} tanındı`} />
        <KpiCard label="Sarf satırı" value={formatNumber(totalRows, 0)} detail={`${formatNumber(matched.length, 0)} odak ürün eşleşti`} />
        <KpiCard
          label="Eşleşme"
          value={matchedRatio === null ? "-" : `%${formatNumber(matchedRatio * 100, 1)}`}
          detail={`${products.length} ürün başlığı taranıyor`}
          tone={matchedRatio && matchedRatio > 0.1 ? "good" : "default"}
        />
        <KpiCard
          label="Doğrulama"
          value={annualTotals.length ? formatNumber(validationProblems, 0) : "-"}
          detail={annualTotals.length ? "fark veren kod satırı" : "yıllık toplam yüklenmedi"}
          tone={validationProblems ? "warn" : "good"}
        />
      </div>

      <section className="panel wide">
        <div className="panel-title">
          <div>
            <h2>Odak ürün kıyası</h2>
            <span>{factories.join(", ") || "Fabrika yok"}</span>
          </div>
          <BarChart3 size={20} />
        </div>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Ürün</th>
                <th>Birim</th>
                {years.map((year) => (
                  <th key={year}>{year}</th>
                ))}
                <th>Toplam</th>
                <th>Satır</th>
              </tr>
            </thead>
            <tbody>
              {productRows.map((row) => (
                <tr key={row.product.product}>
                  <td>
                    <div className="name-cell">
                      <strong>{row.product.product}</strong>
                      <span>{row.product.keywords.slice(0, 3).join(", ")}</span>
                    </div>
                  </td>
                  <td>{row.product.analysisUnit}</td>
                  {row.yearly.map((value, index) => (
                    <td key={`${row.product.product}-${years[index]}`}>{formatNumber(value, 2)}</td>
                  ))}
                  <td>
                    <div className="bar-cell">
                      <span>{formatNumber(row.total, 2)}</span>
                      <i style={{ width: `${Math.max(4, (row.total / maxProductTotal) * 100)}%` }} />
                    </div>
                  </td>
                  <td>{formatNumber(row.rows, 0)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      <section className="panel">
        <div className="panel-title">
          <div>
            <h2>Dosya dağılımı</h2>
            <span>Otomatik sınıflandırma</span>
          </div>
          <Layers3 size={20} />
        </div>
        <div className="stacked-list">
          {(["consumption", "production", "annualTotal", "purchase", "unknown"] as FileKind[]).map((kind) => {
            const count = files.filter((file) => file.kind === kind).length;
            return (
              <div key={kind} className="stacked-row">
                <span>{KIND_LABELS[kind]}</span>
                <strong>{formatNumber(count, 0)}</strong>
              </div>
            );
          })}
        </div>
      </section>

      <section className="panel">
        <div className="panel-title">
          <div>
            <h2>Eşleşmeyen örnekler</h2>
            <span>İlk 12 sarf satırı</span>
          </div>
          <Search size={20} />
        </div>
        <div className="compact-list">
          {records
            .filter((record) => !record.focusProduct)
            .slice(0, 12)
            .map((record) => (
              <div className="compact-row" key={record.id}>
                <span>{record.productName}</span>
                <small>
                  {record.factory} · {formatNumber(record.quantity, 2)} {record.unit}
                </small>
              </div>
            ))}
          {!records.some((record) => !record.focusProduct) && <span className="muted">Eşleşmeyen satır yok.</span>}
        </div>
      </section>
    </div>
  );
}

function ComparePage({
  products,
  records,
  production,
  selectedProduct,
  setSelectedProduct,
  metric,
  setMetric,
  selectedMonth,
  setSelectedMonth,
}: {
  products: FocusProductConfig[];
  records: ClassifiedConsumptionRecord[];
  production: ProductionRecord[];
  selectedProduct: string;
  setSelectedProduct: (product: string) => void;
  metric: Metric;
  setMetric: (metric: Metric) => void;
  selectedMonth: number | null;
  setSelectedMonth: (month: number | null) => void;
}) {
  const years = yearsFromData(records, production);
  const selectedConfig = products.find((product) => product.product === selectedProduct) ?? products[0];
  const filteredFactories = factoriesFromData([
    ...records.filter((record) => record.focusProduct === selectedConfig.product),
    ...production,
  ]);
  const consumptionMap = useMemo(() => consumptionByProductFactoryYearMonth(records), [records]);
  const productionMap = useMemo(() => productionByFactoryYearMonth(production), [production]);

  const matrixRows = filteredFactories.map((factory) => {
    const values = years.map((year) =>
      metricForCell({
        consumptionMap,
        productionMap,
        product: selectedConfig.product,
        factory,
        year,
        month: selectedMonth,
        analysisUnit: selectedConfig.analysisUnit,
        metric,
      }),
    );
    return { factory, values };
  });
  const maxValue = Math.max(...matrixRows.flatMap((row) => row.values.map((value) => Math.abs(value ?? 0))), 1);

  const monthRows = filteredFactories.flatMap((factory) =>
    MONTHS.map((month) => {
      const values = years.map((year) =>
        metricForCell({
          consumptionMap,
          productionMap,
          product: selectedConfig.product,
          factory,
          year,
          month: month.value,
          analysisUnit: selectedConfig.analysisUnit,
          metric,
        }),
      );
      return { factory, month: month.value, values };
    }),
  );

  return (
    <div className="page-grid">
      <section className="panel wide">
        <div className="filter-bar">
          <label>
            <span>Ürün</span>
            <select value={selectedConfig.product} onChange={(event) => setSelectedProduct(event.target.value)}>
              {products.map((product) => (
                <option key={product.product} value={product.product}>
                  {product.product}
                </option>
              ))}
            </select>
          </label>
          <label>
            <span>Ay</span>
            <select value={selectedMonth ?? "all"} onChange={(event) => setSelectedMonth(event.target.value === "all" ? null : Number(event.target.value))}>
              <option value="all">Yıl toplamı</option>
              {MONTHS.map((month) => (
                <option key={month.value} value={month.value}>
                  {month.label}
                </option>
              ))}
            </select>
          </label>
          <div className="segmented" aria-label="Metrik">
            {(["quantity", "ratio", "production", "amount"] as Metric[]).map((item) => (
              <button type="button" key={item} className={metric === item ? "active" : ""} onClick={() => setMetric(item)}>
                {metricLabel(item)}
              </button>
            ))}
          </div>
        </div>
      </section>

      <section className="panel wide">
        <div className="panel-title">
          <div>
            <h2>Fabrika ve yıl kıyası</h2>
            <span>
              {selectedConfig.analysisUnit} · {selectedMonth ? monthLabel(selectedMonth) : "Yıllık"}
            </span>
          </div>
          <GitCompare size={20} />
        </div>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Fabrika</th>
                {years.map((year) => (
                  <th key={year}>{year}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {matrixRows.map((row) => (
                <tr key={row.factory}>
                  <td>
                    <strong>{row.factory}</strong>
                  </td>
                  {row.values.map((value, index) => (
                    <td key={`${row.factory}-${years[index]}`}>
                      <div className="bar-cell">
                        <span>{displayMetric(value, metric)}</span>
                        <i style={{ width: `${Math.max(4, (Math.abs(value ?? 0) / maxValue) * 100)}%` }} />
                      </div>
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      <section className="panel wide">
        <div className="panel-title">
          <div>
            <h2>Aylık kırılım</h2>
            <span>{selectedConfig.product}</span>
          </div>
          <Filter size={20} />
        </div>
        <div className="table-wrap tall">
          <table>
            <thead>
              <tr>
                <th>Fabrika</th>
                <th>Ay</th>
                {years.map((year) => (
                  <th key={year}>{year}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {monthRows
                .filter((row) => row.values.some((value) => Math.abs(value ?? 0) > 0))
                .map((row) => (
                  <tr key={`${row.factory}-${row.month}`}>
                    <td>{row.factory}</td>
                    <td>{monthLabel(row.month)}</td>
                    {row.values.map((value, index) => (
                      <td key={`${row.factory}-${row.month}-${years[index]}`}>{displayMetric(value, metric)}</td>
                    ))}
                  </tr>
                ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

function ValidationPage({
  products,
  records,
  rawConsumption,
  annualTotals,
  purchases,
}: {
  products: FocusProductConfig[];
  records: ClassifiedConsumptionRecord[];
  rawConsumption: ConsumptionRecord[];
  annualTotals: AnnualTotalRecord[];
  purchases: PurchaseRecord[];
}) {
  const exactRows = useMemo(() => buildExactValidationRows(rawConsumption, annualTotals), [rawConsumption, annualTotals]);
  const focusRows = useMemo(() => buildFocusValidationRows(records, annualTotals, products), [records, annualTotals, products]);
  const purchaseRows = useMemo(() => buildPurchaseValidationRows(rawConsumption, purchases), [rawConsumption, purchases]);
  const exactOk = exactRows.filter((row) => row.status === "OK").length;
  const focusOk = focusRows.filter((row) => row.status === "OK").length;

  if (!annualTotals.length && !purchases.length) {
    return <EmptyState title="Doğrulama dosyası bekleniyor" detail="Ürün bazında yıllık toplam veya satın alma dosyaları yüklendiğinde mutabakat burada açılır." />;
  }

  return (
    <div className="page-grid">
      <div className="kpi-grid">
        <KpiCard label="Kod bazlı" value={`${formatNumber(exactOk, 0)} / ${formatNumber(exactRows.length, 0)}`} detail="yıllık toplam mutabakatı" tone={exactOk === exactRows.length ? "good" : "warn"} />
        <KpiCard label="Odak ürün" value={`${formatNumber(focusOk, 0)} / ${formatNumber(focusRows.length, 0)}`} detail="keyword grubu mutabakatı" tone={focusOk === focusRows.length ? "good" : "warn"} />
        <KpiCard label="Satın alma" value={formatNumber(purchases.length, 0)} detail="yüklenen satın alma satırı" />
      </div>

      <section className="panel wide">
        <div className="panel-title">
          <div>
            <h2>Kod bazlı doğrulama</h2>
            <span>Satır bazlı sarf toplamı ile yıllık ürün toplamı</span>
          </div>
          <ShieldCheck size={20} />
        </div>
        <div className="table-wrap tall">
          <table>
            <thead>
              <tr>
                <th>Durum</th>
                <th>Yıl</th>
                <th>Fabrika</th>
                <th>Ürün kodu</th>
                <th>Ürün adı</th>
                <th>Birim</th>
                <th>Satır toplam</th>
                <th>Yıllık toplam</th>
                <th>Fark</th>
              </tr>
            </thead>
            <tbody>
              {exactRows.slice(0, 200).map((row) => (
                <tr key={row.key}>
                  <td>
                    <span className={statusClass(row.status)}>{row.status}</span>
                  </td>
                  <td>{row.year ?? "-"}</td>
                  <td>{row.factory}</td>
                  <td>{row.productCode}</td>
                  <td className="clip">{row.productName}</td>
                  <td>{row.unit}</td>
                  <td>{formatNumber(row.detailQuantity, 2)}</td>
                  <td>{formatNumber(row.annualQuantity, 2)}</td>
                  <td>{formatNumber(row.diff, 2)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      <section className="panel wide">
        <div className="panel-title">
          <div>
            <h2>Odak ürün doğrulama</h2>
            <span>Keyword ile sınıflandırılmış ürün grupları</span>
          </div>
          <CheckCircle2 size={20} />
        </div>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Durum</th>
                <th>Yıl</th>
                <th>Fabrika</th>
                <th>Ürün</th>
                <th>Birim</th>
                <th>Satır toplam</th>
                <th>Yıllık toplam</th>
                <th>Fark</th>
              </tr>
            </thead>
            <tbody>
              {focusRows.slice(0, 120).map((row) => (
                <tr key={row.key}>
                  <td>
                    <span className={statusClass(row.status)}>{row.status}</span>
                  </td>
                  <td>{row.year ?? "-"}</td>
                  <td>{row.factory}</td>
                  <td>{row.product}</td>
                  <td>{row.unit}</td>
                  <td>{formatNumber(row.detailQuantity, 2)}</td>
                  <td>{formatNumber(row.annualQuantity, 2)}</td>
                  <td>{formatNumber(row.diff, 2)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      {purchases.length > 0 && (
        <section className="panel wide">
          <div className="panel-title">
            <div>
              <h2>Satın alma eşleşmesi</h2>
              <span>Ürün kodu ve yıl toplamı</span>
            </div>
            <FolderInput size={20} />
          </div>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Durum</th>
                  <th>Yıl</th>
                  <th>Ürün kodu</th>
                  <th>Ürün adı</th>
                  <th>Birim</th>
                  <th>Sarf</th>
                  <th>Satın alma</th>
                  <th>Fark</th>
                </tr>
              </thead>
              <tbody>
                {purchaseRows.slice(0, 160).map((row) => (
                  <tr key={row.key}>
                    <td>
                      <span className={statusClass(row.status)}>{row.status}</span>
                    </td>
                    <td>{row.year ?? "-"}</td>
                    <td>{row.productCode}</td>
                    <td className="clip">{row.productName}</td>
                    <td>{row.unit}</td>
                    <td>{formatNumber(row.consumptionQuantity, 2)}</td>
                    <td>{formatNumber(row.purchaseQuantity, 2)}</td>
                    <td>{formatNumber(row.diff, 2)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}
    </div>
  );
}

function DataPage({
  products,
  setProducts,
  resetProducts,
  files,
  records,
}: {
  products: FocusProductConfig[];
  setProducts: (products: FocusProductConfig[]) => void;
  resetProducts: () => void;
  files: ParsedFileSummary[];
  records: ClassifiedConsumptionRecord[];
}) {
  function updateProduct(productName: string, updater: (item: FocusProductConfig) => FocusProductConfig) {
    setProducts(products.map((item) => (item.product === productName ? updater(item) : item)));
  }

  return (
    <div className="page-grid">
      <section className="panel wide">
        <div className="panel-title">
          <div>
            <h2>Yüklenen dosyalar</h2>
            <span>{formatNumber(files.length, 0)} dosya</span>
          </div>
          <Database size={20} />
        </div>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Tür</th>
                <th>Dosya</th>
                <th>Sayfa</th>
                <th>Satır</th>
                <th>Boyut</th>
                <th>Uyarı</th>
              </tr>
            </thead>
            <tbody>
              {files.map((file) => (
                <tr key={file.key}>
                  <td>
                    <span className={`kind ${file.kind}`}>{KIND_LABELS[file.kind]}</span>
                  </td>
                  <td className="clip">{file.path}</td>
                  <td>{file.sheets.join(", ")}</td>
                  <td>{formatNumber(file.rows, 0)}</td>
                  <td>{fileSizeLabel(file.size)}</td>
                  <td className="clip">{file.warnings.join(" · ")}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      <section className="panel wide">
        <div className="panel-title">
          <div>
            <h2>Ürün sözlüğü</h2>
            <span>Malzeme adında aranacak kelimeler</span>
          </div>
          <button type="button" className="icon-button" onClick={resetProducts} title="Varsayılan listeye dön">
            <RefreshCw size={18} />
          </button>
        </div>
        <div className="keyword-grid">
          {products.map((product) => (
            <div className="keyword-row" key={product.product}>
              <strong>{product.product}</strong>
              <select
                value={product.analysisUnit}
                onChange={(event) =>
                  updateProduct(product.product, (item) => ({ ...item, analysisUnit: event.target.value as FocusProductConfig["analysisUnit"] }))
                }
              >
                {ANALYSIS_UNITS.map((unit) => (
                  <option key={unit} value={unit}>
                    {unit}
                  </option>
                ))}
              </select>
              <input
                value={product.keywords.join(", ")}
                onChange={(event) =>
                  updateProduct(product.product, (item) => ({
                    ...item,
                    keywords: event.target.value
                      .split(",")
                      .map((keyword) => keyword.trim())
                      .filter(Boolean),
                  }))
                }
              />
            </div>
          ))}
        </div>
      </section>

      <section className="panel wide">
        <div className="panel-title">
          <div>
            <h2>Dışa aktarım</h2>
            <span>Sınıflandırılmış satır verisi</span>
          </div>
          <Download size={20} />
        </div>
        <div className="action-row">
          <button type="button" className="button" onClick={() => downloadCsv("siniflandirilmis-sarf.csv", exportRowsForConsumption(records))} disabled={!records.length}>
            <Download size={17} />
            CSV indir
          </button>
        </div>
      </section>
    </div>
  );
}

export default function App() {
  const [activeTab, setActiveTab] = useState<Tab>("overview");
  const [products, setProducts] = useState<FocusProductConfig[]>(loadProducts);
  const [files, setFiles] = useState<ParsedFileSummary[]>([]);
  const [consumption, setConsumption] = useState<ConsumptionRecord[]>([]);
  const [production, setProduction] = useState<ProductionRecord[]>([]);
  const [annualTotals, setAnnualTotals] = useState<AnnualTotalRecord[]>([]);
  const [purchases, setPurchases] = useState<PurchaseRecord[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [progress, setProgress] = useState("");
  const [selectedProduct, setSelectedProduct] = useState(DEFAULT_PRODUCTS[0].product);
  const [metric, setMetric] = useState<Metric>("quantity");
  const [selectedMonth, setSelectedMonth] = useState<number | null>(null);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(products));
  }, [products]);

  const classifiedRecords = useMemo(() => classifyConsumption(consumption, products), [consumption, products]);

  useEffect(() => {
    if (!products.some((product) => product.product === selectedProduct)) {
      setSelectedProduct(products[0]?.product ?? "");
    }
  }, [products, selectedProduct]);

  async function handleFiles(fileList: FileList | null) {
    if (!fileList?.length) return;
    const existing = new Set(files.map((file) => file.key));
    const incoming = Array.from(fileList).filter((file) => !existing.has(makeFileKey(file)));
    if (!incoming.length) return;

    setIsLoading(true);
    try {
      const parsed = await parseWorkbookFiles(incoming, setProgress);
      setFiles((current) => current.concat(parsed.files));
      setConsumption((current) => current.concat(parsed.consumption));
      setProduction((current) => current.concat(parsed.production));
      setAnnualTotals((current) => current.concat(parsed.annualTotals));
      setPurchases((current) => current.concat(parsed.purchases));
      setProgress("");
    } finally {
      setIsLoading(false);
    }
  }

  function clearAll() {
    setFiles([]);
    setConsumption([]);
    setProduction([]);
    setAnnualTotals([]);
    setPurchases([]);
    setProgress("");
  }

  function resetProducts() {
    setProducts(DEFAULT_PRODUCTS);
  }

  const tabs: Array<{ id: Tab; label: string; icon: typeof BarChart3 }> = [
    { id: "overview", label: "Ana kıyas", icon: BarChart3 },
    { id: "compare", label: "Yıl / fabrika", icon: GitCompare },
    { id: "validation", label: "Doğrulama", icon: ShieldCheck },
    { id: "data", label: "Veri & sözlük", icon: Settings2 },
  ];

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark">
            <Layers3 size={24} />
          </div>
          <div>
            <h1>Sarf Analiz Masası</h1>
            <span>Satır bazlı sarf, üretim ve yıllık toplam mutabakatı</span>
          </div>
        </div>
        <div className="top-actions">
          <button type="button" className="icon-button" onClick={clearAll} title="Veriyi temizle" disabled={!files.length || isLoading}>
            <Trash2 size={18} />
          </button>
        </div>
      </header>

      <main>
        <UploadPanel onFiles={handleFiles} isLoading={isLoading} progress={progress} />

        <nav className="tabs" aria-label="Sayfalar">
          {tabs.map((tab) => {
            const Icon = tab.icon;
            return (
              <button type="button" key={tab.id} className={activeTab === tab.id ? "active" : ""} onClick={() => setActiveTab(tab.id)}>
                <Icon size={17} />
                {tab.label}
              </button>
            );
          })}
        </nav>

        {files.some((file) => file.kind === "unknown") && (
          <div className="notice">
            <AlertTriangle size={18} />
            <span>Bazı dosyalar tanınmadı. Veri & sözlük sayfasındaki uyarılardan kolon adlarını kontrol edebilirsin.</span>
          </div>
        )}

        {activeTab === "overview" && (
          <OverviewPage products={products} records={classifiedRecords} production={production} annualTotals={annualTotals} files={files} />
        )}
        {activeTab === "compare" && (
          <ComparePage
            products={products}
            records={classifiedRecords}
            production={production}
            selectedProduct={selectedProduct}
            setSelectedProduct={setSelectedProduct}
            metric={metric}
            setMetric={setMetric}
            selectedMonth={selectedMonth}
            setSelectedMonth={setSelectedMonth}
          />
        )}
        {activeTab === "validation" && (
          <ValidationPage products={products} records={classifiedRecords} rawConsumption={consumption} annualTotals={annualTotals} purchases={purchases} />
        )}
        {activeTab === "data" && <DataPage products={products} setProducts={setProducts} resetProducts={resetProducts} files={files} records={classifiedRecords} />}
      </main>
    </div>
  );
}
