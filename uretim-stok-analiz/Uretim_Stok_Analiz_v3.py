#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Üretim ve Stok Analiz Paneli  v3
· Üretim Raporları  → Grafik 1-3 (üretim trendi, oran, stok)
· Sevk Raporları    → Grafik 4-6 (sevk trendi, üretim-sevk, stok-sevk)
· MA2 / MA6 hareketli ortalama toggleları
· Gün filtresi: sağ panelde tarih checkbox listesi
Gereksinimler: pip install pdfplumber matplotlib
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re, math, threading
from datetime import date, datetime
from pathlib import Path
from typing import Optional

import pdfplumber
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import matplotlib.dates as mdates
import matplotlib.ticker as mticker
from matplotlib import rcParams

# ── Sabitler ──────────────────────────────────────────────────────────────────

FACTORIES = ["Tesis 1", "Tesis 2", "Tesis 3", "Tesis 4", "Tesis 5"]

FAB_COLORS = {
    "Tesis 1": "#E74C3C",
    "Tesis 2":     "#3498DB",
    "Tesis 3":     "#2ECC71",
    "Tesis 4":     "#F39C12",
    "Tesis 5":     "#9B59B6",
}

C1_METRICS = {
    "toplam":  ("Toplam Üretim",  "-",       "o"),
    "kal1":    ("1.Kalite",        "--",      "s"),
    "export":  ("Export (Defolu)", "-.",      "^"),
    "iskarta": ("Iskarta",         ":",       "D"),
    "stok":    ("Toplam Stok",     (0,(4,2)), "x"),
}

C2_RATIOS = {
    "kal1_pct":    ("1.Kalite %",  "-",  "o"),
    "export_pct":  ("Export %",    "--", "^"),
    "iskarta_pct": ("Iskarta %",   ":",  "D"),
}

SEVK_METRICS = {
    "ic_piyasa":   ("İç Piyasa",   "-",  "o"),
    "ihracat":     ("İhracat",     "--", "^"),
    "toplam_sevk": ("Toplam Sevk", "-.", "s"),
}

DATE_RE  = re.compile(r"(\d{2})-(\d{2})-(\d{4})")
_TARGET  = set(FACTORIES)

rcParams["font.family"] = "sans-serif"

# ── PDF: Üretim & Stok ────────────────────────────────────────────────────────

def _num(s) -> float:
    if not s:
        return float("nan")
    s = str(s).strip()
    if s in ("", "-", "0,0"):
        return 0.0
    try:
        return float(s.replace(".", "").replace(",", "."))
    except ValueError:
        return float("nan")


def extract_uretim_pdf(path: Path) -> Optional[dict]:
    try:
        with pdfplumber.open(str(path)) as pdf:
            result = {}
            t1 = pdf.pages[0].extract_tables()
            if t1:
                for row in t1[0][2:]:
                    if not row or not row[0]:
                        continue
                    fab = row[0].strip()
                    if fab not in _TARGET:
                        continue
                    result[fab] = {
                        "kal1":    _num(row[1]  if len(row) > 1  else ""),
                        "export":  _num(row[6]  if len(row) > 6  else ""),
                        "iskarta": _num(row[8]  if len(row) > 8  else ""),
                        "toplam":  _num(row[10] if len(row) > 10 else ""),
                        "stok":    float("nan"),
                    }
            if len(pdf.pages) > 2:
                t3 = pdf.pages[2].extract_tables()
                if t3:
                    for row in t3[0][1:]:
                        if not row or len(row) < 3:
                            continue
                        fab = (row[1] or "").strip()
                        if fab in _TARGET:
                            result.setdefault(fab, {
                                "kal1": float("nan"), "export": float("nan"),
                                "iskarta": float("nan"), "toplam": float("nan"),
                            })["stok"] = _num(row[2])
            return result or None
    except Exception as e:
        print(f"  ✗ Üretim PDF {path.name}: {e}")
        return None


def extract_sevk_pdf(path: Path) -> Optional[dict]:
    try:
        with pdfplumber.open(str(path)) as pdf:
            result = {}
            tables = pdf.pages[0].extract_tables()
            if not tables:
                return None
            for row in tables[0][3:]:
                if not row or not row[0]:
                    continue
                fab = row[0].strip()
                if fab == "Toplam":
                    break
                if fab not in _TARGET:
                    continue
                result[fab] = {
                    "ic_piyasa":   _num(row[3]  if len(row) > 3  else ""),
                    "ihracat":     _num(row[8]  if len(row) > 8  else ""),
                    "toplam_sevk": _num(row[13] if len(row) > 13 else ""),
                }
            return result or None
    except Exception as e:
        print(f"  ✗ Sevk PDF {path.name}: {e}")
        return None


def date_from_path(path: Path) -> Optional[date]:
    m = DATE_RE.search(path.stem)
    if not m:
        return None
    try:
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


# ── DatePanel widget ──────────────────────────────────────────────────────────

class DatePanel(ttk.LabelFrame):
    """Sağ taraftaki kaydırılabilir tarih filtre paneli."""

    def __init__(self, parent, callback, **kw):
        kw.setdefault("text", "  Günler  ")
        kw.setdefault("padding", (4, 4))
        super().__init__(parent, **kw)

        btn_row = tk.Frame(self, bg="#FFFFFF")
        btn_row.pack(fill="x", padx=2, pady=(2, 2))
        ttk.Button(btn_row, text="Tümü",    width=5,
                   command=lambda: self._set_all(True)).pack(side="left", padx=1)
        ttk.Button(btn_row, text="Hiçbiri", width=7,
                   command=lambda: self._set_all(False)).pack(side="left", padx=1)

        container = tk.Frame(self, bg="#FFFFFF")
        container.pack(fill="both", expand=True)

        self._cv = tk.Canvas(container, bg="#FFFFFF", highlightthickness=0, width=118)
        sb = ttk.Scrollbar(container, orient="vertical", command=self._cv.yview)
        self._cv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._cv.pack(side="left", fill="both", expand=True)

        self._inner = tk.Frame(self._cv, bg="#FFFFFF")
        self._win   = self._cv.create_window((0, 0), window=self._inner, anchor="nw")

        self._inner.bind("<Configure>", lambda e: self._cv.configure(
            scrollregion=self._cv.bbox("all")))
        self._cv.bind("<Configure>", lambda e: self._cv.itemconfig(
            self._win, width=e.width))

        self._vars: dict[date, tk.BooleanVar] = {}
        self._callback = callback

    def populate(self, dates: list):
        for w in self._inner.winfo_children():
            w.destroy()
        self._vars.clear()
        for d in sorted(dates):
            var = tk.BooleanVar(value=True)
            self._vars[d] = var
            tk.Checkbutton(
                self._inner, text=d.strftime("%d.%m.%Y"),
                variable=var, command=self._callback,
                bg="#FFFFFF", font=("Helvetica", 9), anchor="w",
            ).pack(fill="x", padx=2)
        self._cv.configure(scrollregion=self._cv.bbox("all"))

    def active_dates(self) -> list:
        return sorted(d for d, v in self._vars.items() if v.get())

    def _set_all(self, val: bool):
        for v in self._vars.values():
            v.set(val)
        self._callback()


# ── DateEntry widget ──────────────────────────────────────────────────────────

class DateEntry(tk.Frame):
    def __init__(self, parent, initial: date = None, **kw):
        super().__init__(parent, **kw)
        d = initial or date.today()
        self.day_var   = tk.StringVar(value=f"{d.day:02d}")
        self.month_var = tk.StringVar(value=f"{d.month:02d}")
        self.year_var  = tk.StringVar(value=str(d.year))
        vcmd = (self.register(lambda v: v == "" or v.isdigit()), "%P")
        for var, w, sep in [(self.day_var, 3, True), (self.month_var, 3, True),
                             (self.year_var, 5, False)]:
            ttk.Entry(self, textvariable=var, width=w, justify="center",
                      validate="key", validatecommand=vcmd).pack(side="left")
            if sep:
                ttk.Label(self, text=".").pack(side="left")

    def get(self) -> date:
        try:
            return date(int(self.year_var.get()),
                        int(self.month_var.get()),
                        int(self.day_var.get()))
        except (ValueError, TypeError) as e:
            raise ValueError(f"Geçersiz tarih: {e}")


# ── Ana Uygulama ──────────────────────────────────────────────────────────────

class App(tk.Tk):
    BG   = "#F0F4F8"
    CARD = "#FFFFFF"

    def __init__(self):
        super().__init__()
        self.title("Üretim ve Stok Analiz Paneli")
        self.configure(bg=self.BG)
        self.minsize(1280, 760)
        self.resizable(True, True)

        # Veri depoları
        self._uretim_data: dict[date, dict] = {}
        self._sevk_data:   dict[date, dict] = {}
        self._uretim_files: list[Path] = []
        self._sevk_files:   list[Path] = []

        # Fabrika toggle'ları
        self._fab_vars = {f: tk.BooleanVar(value=True) for f in FACTORIES}

        # Grafik 1 metrikleri
        self._c1_vars = {k: tk.BooleanVar(value=(k in {"toplam", "kal1"}))
                         for k in C1_METRICS}
        # Grafik 2 oranları
        self._c2_vars = {k: tk.BooleanVar(value=True) for k in C2_RATIOS}
        # Grafik 3
        self._c3_abs_var   = tk.BooleanVar(value=True)
        self._c3_pct_var   = tk.BooleanVar(value=False)
        self._c3_total_var = tk.BooleanVar(value=True)
        # Grafik 4 sevk metrikleri
        self._c4_vars      = {k: tk.BooleanVar(value=True) for k in SEVK_METRICS}
        self._c4_total_var = tk.BooleanVar(value=True)
        # Grafik 5
        self._c5_uretim_var = tk.StringVar(value="toplam")
        self._c5_sevk_var   = tk.StringVar(value="toplam_sevk")
        # Grafik 6
        self._c6_sevk_var   = tk.StringVar(value="toplam_sevk")

        # MA toggle'ları: [idx][0]=MA2, [idx][1]=MA6
        self._ma_vars = [
            [tk.BooleanVar(value=False), tk.BooleanVar(value=False)]
            for _ in range(6)
        ]

        # Figürler ve gün panelleri
        self._figs       = [None] * 6
        self._axes       = [None] * 6
        self._ax3r       = None
        self._canvases   = [None] * 6
        self._date_panels: list[Optional[DatePanel]] = [None] * 6

        self._style_setup()
        self._build_ui()

    # ── Stil ─────────────────────────────────────────────────────────────────

    def _style_setup(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame",      background=self.BG)
        s.configure("TLabel",      background=self.BG, font=("Helvetica", 11))
        s.configure("TLabelframe", background=self.CARD, relief="solid",
                    borderwidth=1, bordercolor="#D1D9E0")
        s.configure("TLabelframe.Label", background=self.CARD,
                    foreground="#1A73E8", font=("Helvetica", 11, "bold"))
        s.configure("TCheckbutton", background=self.CARD, font=("Helvetica", 10))
        s.configure("TButton",      font=("Helvetica", 11), padding=(8, 5))
        s.configure("Big.TButton",  font=("Helvetica", 12, "bold"), padding=(10, 8))
        s.configure("TEntry",       fieldbackground="white", font=("Helvetica", 11))
        s.configure("TNotebook",    background=self.BG)
        s.configure("TNotebook.Tab", font=("Helvetica", 11), padding=(10, 6))
        s.configure("TRadiobutton", background=self.CARD, font=("Helvetica", 10))

    # ── UI İnşası ────────────────────────────────────────────────────────────

    def _build_ui(self):
        left = ttk.Frame(self, padding=(12, 12, 6, 12))
        left.pack(side="left", fill="y")

        right = ttk.Frame(self, padding=(6, 12, 12, 12))
        right.pack(side="left", fill="both", expand=True)

        self._build_left(left)
        self._build_notebook(right)

        self._status_var = tk.StringVar(value="Dosya yüklenmedi.")
        ttk.Label(self, textvariable=self._status_var,
                  foreground="#718096", background=self.BG,
                  font=("Helvetica", 10)).pack(side="bottom", anchor="w",
                                               padx=12, pady=4)

    # ── Sol Panel ────────────────────────────────────────────────────────────

    def _build_left(self, parent):
        ttk.Label(parent, text="Analiz Paneli",
                  font=("Helvetica", 14, "bold"),
                  foreground="#1A73E8",
                  background=self.BG).pack(anchor="w", pady=(0, 10))

        self._lb_uretim, _ = self._file_section(
            parent, "  📊 Üretim Raporları  ",
            "(Uretim ve Stok Raporu)",
            self._add_uretim_files, self._remove_uretim_file, self._clear_uretim_files,
        )

        self._lb_sevk, _ = self._file_section(
            parent, "  🚚 Günlük Sevk Raporları  ",
            "(Gunluk Sevk Raporu)",
            self._add_sevk_files, self._remove_sevk_file, self._clear_sevk_files,
        )

        df = ttk.LabelFrame(parent, text="  Tarih Aralığı  ", padding=(8, 6))
        df.pack(fill="x", pady=(0, 10))
        for row_i, (lbl, attr, init) in enumerate([
            ("Başlangıç:", "_start_e", date.today().replace(day=1)),
            ("Bitiş:",     "_end_e",   date.today()),
        ]):
            ttk.Label(df, text=lbl, background=self.CARD,
                      font=("Helvetica", 10)).grid(row=row_i, column=0, sticky="w")
            entry = DateEntry(df, initial=init)
            entry.configure(background=self.CARD)
            entry.grid(row=row_i, column=1, sticky="w", pady=2)
            setattr(self, attr, entry)

        fabf = ttk.LabelFrame(parent, text="  Fabrikalara  ", padding=(8, 6))
        fabf.pack(fill="x", pady=(0, 10))
        for fab in FACTORIES:
            row = ttk.Frame(fabf)
            row.pack(anchor="w", pady=1)
            dot = tk.Canvas(row, width=12, height=12, bg=self.CARD,
                            highlightthickness=0)
            dot.create_rectangle(1, 1, 11, 11, fill=FAB_COLORS[fab], outline="")
            dot.pack(side="left", padx=(0, 4))
            ttk.Checkbutton(row, text=fab, variable=self._fab_vars[fab],
                            command=self._on_toggle).pack(side="left")

        ttk.Button(parent, text="▶  Grafikle",
                   command=self._load_and_plot,
                   style="Big.TButton").pack(fill="x", pady=(4, 0))

    def _file_section(self, parent, title, hint, add_cmd, rem_cmd, clr_cmd):
        frame = ttk.LabelFrame(parent, text=title, padding=(8, 6))
        frame.pack(fill="x", pady=(0, 8))
        ttk.Label(frame, text=hint, background=self.CARD,
                  foreground="#718096", font=("Helvetica", 9)).pack(anchor="w")
        lb = tk.Listbox(frame, height=5, selectmode="extended",
                        font=("Helvetica", 9), bg="white",
                        relief="flat", borderwidth=1)
        lb.pack(fill="x", pady=(2, 4))
        btn_row = ttk.Frame(frame)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="+ Ekle",  command=add_cmd).pack(side="left", padx=(0, 4))
        ttk.Button(btn_row, text="Sil",     command=rem_cmd).pack(side="left", padx=(0, 4))
        ttk.Button(btn_row, text="Temizle", command=clr_cmd).pack(side="left")
        return lb, frame

    # ── Dosya yönetimi ────────────────────────────────────────────────────────

    def _add_uretim_files(self):
        self._add_files_to(self._lb_uretim, self._uretim_files,
                           "Üretim ve Stok Raporu PDF")

    def _add_sevk_files(self):
        self._add_files_to(self._lb_sevk, self._sevk_files,
                           "Günlük Sevk Raporu PDF")

    def _add_files_to(self, lb, lst, title):
        paths = filedialog.askopenfilenames(
            title=f"{title} Seçin",
            filetypes=[("PDF", "*.pdf"), ("Tümü", "*.*")],
        )
        for p in paths:
            pp = Path(p)
            if pp not in lst:
                lst.append(pp)
                lb.insert("end", pp.name)

    def _remove_uretim_file(self):
        self._remove_from(self._lb_uretim, self._uretim_files)

    def _remove_sevk_file(self):
        self._remove_from(self._lb_sevk, self._sevk_files)

    def _remove_from(self, lb, lst):
        for i in reversed(lb.curselection()):
            lb.delete(i)
            lst.pop(i)

    def _clear_uretim_files(self):
        self._lb_uretim.delete(0, "end")
        self._uretim_files.clear()

    def _clear_sevk_files(self):
        self._lb_sevk.delete(0, "end")
        self._sevk_files.clear()

    # ── Notebook ─────────────────────────────────────────────────────────────

    def _build_notebook(self, parent):
        nb = ttk.Notebook(parent)
        nb.pack(fill="both", expand=True)

        specs = [
            ("📊  Üretim Trendi",   self._build_tab1),
            ("📈  Oran Trendi",     self._build_tab2),
            ("🏭  Stok Trendi",     self._build_tab3),
            ("🚚  Sevk Trendi",     self._build_tab4),
            ("⚖️  Üretim vs Sevk", self._build_tab5),
            ("📦  Stok vs Sevk",   self._build_tab6),
        ]
        for label, builder in specs:
            tab = ttk.Frame(nb)
            nb.add(tab, text=label)
            builder(tab)

    # ── Grafik yardımcıları ───────────────────────────────────────────────────

    def _make_chart(self, parent, idx):
        fig  = Figure(figsize=(10, 5.5), dpi=100, facecolor=self.CARD)
        ax   = fig.add_subplot(111)
        self._figs[idx] = fig
        self._axes[idx] = ax
        canvas = FigureCanvasTkAgg(fig, parent)
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self._canvases[idx] = canvas
        NavigationToolbar2Tk(canvas, parent)
        return fig, ax, canvas

    def _ctrl_row(self, parent, label="Göster:"):
        ctrl = tk.Frame(parent, bg=self.BG, pady=4)
        ctrl.pack(fill="x", padx=8)
        tk.Label(ctrl, text=label, bg=self.BG,
                 font=("Helvetica", 10, "bold")).pack(side="left", padx=(0, 8))
        return ctrl

    def _add_ma_and_content(self, parent, idx, draw_fn):
        """MA toggles satırı + (DatePanel sağda | grafik solda) içerik alanı."""
        ma_ctrl = self._ctrl_row(parent, "Hareketli Ort.:")
        ttk.Checkbutton(ma_ctrl, text="MA2 (2 günlük)",
                        variable=self._ma_vars[idx][0],
                        command=draw_fn).pack(side="left", padx=4)
        ttk.Checkbutton(ma_ctrl, text="MA6 (6 günlük)",
                        variable=self._ma_vars[idx][1],
                        command=draw_fn).pack(side="left", padx=4)

        content = tk.Frame(parent, bg=self.BG)
        content.pack(fill="both", expand=True)

        dp = DatePanel(content, draw_fn)
        dp.pack(side="right", fill="y", padx=(4, 0), pady=4)
        self._date_panels[idx] = dp

        chart_frame = tk.Frame(content, bg=self.CARD)
        chart_frame.pack(side="left", fill="both", expand=True)
        self._make_chart(chart_frame, idx)

    @staticmethod
    def _fmt_axes(ax, dates):
        n = len(dates)
        ax.xaxis.set_major_formatter(
            mdates.DateFormatter("%d.%m" if n <= 60 else "%m.%y"))
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        ax.figure.autofmt_xdate(rotation=35, ha="right")
        ax.yaxis.set_major_formatter(
            mticker.FuncFormatter(lambda v, _: f"{int(v):,}".replace(",", ".")))
        ax.grid(True, alpha=0.25, linestyle="--")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

    @staticmethod
    def _legend(ax, handles=None, labels=None):
        if handles is None:
            handles, labels = ax.get_legend_handles_labels()
        if handles:
            ax.legend(handles, labels, loc="best",
                      ncol=2, fontsize=8, framealpha=0.9)

    @staticmethod
    def _to_dt(dates):
        return [datetime(d.year, d.month, d.day) for d in dates]

    @staticmethod
    def _ma(values: list, window: int) -> list:
        """NaN-aware hareketli ortalama. NaN değerde buffer sıfırlanır."""
        out = [float("nan")] * len(values)
        buf: list[float] = []
        for i, v in enumerate(values):
            if math.isnan(v):
                buf = []
                continue
            buf.append(v)
            if len(buf) >= window:
                out[i] = sum(buf[-window:]) / window
        return out

    def _filtered_dates(self, idx: int, source_fn) -> list:
        """source_fn sonucunu DatePanel seçimine göre filtreler."""
        all_dates = source_fn()
        dp = self._date_panels[idx]
        if dp is None or not dp._vars:
            return all_dates
        active = set(dp.active_dates())
        return [d for d in all_dates if d in active]

    # ── Tab inşaları ─────────────────────────────────────────────────────────

    def _build_tab1(self, p):
        ctrl = self._ctrl_row(p)
        for k, (lbl, *_) in C1_METRICS.items():
            ttk.Checkbutton(ctrl, text=lbl, variable=self._c1_vars[k],
                            command=self._draw_chart1).pack(side="left", padx=4)
        self._add_ma_and_content(p, 0, self._draw_chart1)

    def _build_tab2(self, p):
        ctrl = self._ctrl_row(p)
        for k, (lbl, *_) in C2_RATIOS.items():
            ttk.Checkbutton(ctrl, text=lbl, variable=self._c2_vars[k],
                            command=self._draw_chart2).pack(side="left", padx=4)
        self._add_ma_and_content(p, 1, self._draw_chart2)

    def _build_tab3(self, p):
        ctrl = self._ctrl_row(p)
        for var, lbl in [(self._c3_abs_var,   "Miktarsal (m²)"),
                         (self._c3_pct_var,   "Oran (% toplam stok)"),
                         (self._c3_total_var, "Genel Toplam Stok")]:
            ttk.Checkbutton(ctrl, text=lbl, variable=var,
                            command=self._draw_chart3).pack(side="left", padx=4)
        self._add_ma_and_content(p, 2, self._draw_chart3)

    def _build_tab4(self, p):
        ctrl = self._ctrl_row(p)
        for k, (lbl, *_) in SEVK_METRICS.items():
            ttk.Checkbutton(ctrl, text=lbl, variable=self._c4_vars[k],
                            command=self._draw_chart4).pack(side="left", padx=4)
        ttk.Separator(ctrl, orient="vertical").pack(side="left", fill="y", padx=8, pady=2)
        ttk.Checkbutton(ctrl, text="Günlük Genel Toplam",
                        variable=self._c4_total_var,
                        command=self._draw_chart4).pack(side="left", padx=4)
        self._add_ma_and_content(p, 3, self._draw_chart4)

    def _build_tab5(self, p):
        ctrl1 = self._ctrl_row(p, "Üretim:")
        for k, (lbl, *_) in C1_METRICS.items():
            if k == "stok":
                continue
            ttk.Radiobutton(ctrl1, text=lbl, value=k,
                            variable=self._c5_uretim_var,
                            command=self._draw_chart5).pack(side="left", padx=4)

        ctrl2 = self._ctrl_row(p, "Sevk:")
        for k, (lbl, *_) in SEVK_METRICS.items():
            ttk.Radiobutton(ctrl2, text=lbl, value=k,
                            variable=self._c5_sevk_var,
                            command=self._draw_chart5).pack(side="left", padx=4)

        self._add_ma_and_content(p, 4, self._draw_chart5)

    def _build_tab6(self, p):
        ctrl = self._ctrl_row(p, "Sevk türü:")
        for k, (lbl, *_) in SEVK_METRICS.items():
            ttk.Radiobutton(ctrl, text=lbl, value=k,
                            variable=self._c6_sevk_var,
                            command=self._draw_chart6).pack(side="left", padx=4)
        self._add_ma_and_content(p, 5, self._draw_chart6)

    # ── Veri yükleme ─────────────────────────────────────────────────────────

    def _load_and_plot(self):
        if not self._uretim_files and not self._sevk_files:
            messagebox.showwarning("Uyarı", "Lütfen en az bir PDF dosyası ekleyin.")
            return
        try:
            d_start = self._start_e.get()
            d_end   = self._end_e.get()
        except ValueError as e:
            messagebox.showerror("Tarih Hatası", str(e))
            return
        if d_start > d_end:
            messagebox.showerror("Tarih Hatası",
                                  "Başlangıç tarihi bitiş tarihinden sonra olamaz.")
            return

        self._status_var.set("Dosyalar okunuyor…")
        self.update_idletasks()

        def worker():
            uretim, sevk = {}, {}
            u_ok = u_skip = u_err = 0
            s_ok = s_skip = s_err = 0

            for path in self._uretim_files:
                d = date_from_path(path)
                if d is None or not (d_start <= d <= d_end):
                    u_skip += 1; continue
                r = extract_uretim_pdf(path)
                if r:
                    uretim[d] = r; u_ok += 1
                else:
                    u_err += 1

            for path in self._sevk_files:
                d = date_from_path(path)
                if d is None or not (d_start <= d <= d_end):
                    s_skip += 1; continue
                r = extract_sevk_pdf(path)
                if r:
                    sevk[d] = r; s_ok += 1
                else:
                    s_err += 1

            self.after(0, self._on_loaded, uretim, sevk,
                       u_ok, u_skip, u_err, s_ok, s_skip, s_err)

        threading.Thread(target=worker, daemon=True).start()

    def _on_loaded(self, uretim, sevk, u_ok, u_skip, u_err, s_ok, s_skip, s_err):
        self._uretim_data = uretim
        self._sevk_data   = sevk

        # Gün panellerini doldur
        u_dates   = sorted(uretim)
        s_dates   = sorted(sevk)
        all_dates = sorted(set(u_dates) | set(s_dates))
        for idx in (0, 1, 2):
            if self._date_panels[idx]:
                self._date_panels[idx].populate(u_dates)
        if self._date_panels[3]:
            self._date_panels[3].populate(s_dates)
        for idx in (4, 5):
            if self._date_panels[idx]:
                self._date_panels[idx].populate(all_dates)

        parts = []
        if uretim:
            ud = sorted(uretim)
            parts.append(f"Üretim: {u_ok} dosya ({ud[0].strftime('%d.%m')}–{ud[-1].strftime('%d.%m.%Y')})")
        if u_err:
            parts.append(f"{u_err} hata")
        if sevk:
            sd = sorted(sevk)
            parts.append(f"Sevk: {s_ok} dosya ({sd[0].strftime('%d.%m')}–{sd[-1].strftime('%d.%m.%Y')})")
        if s_err:
            parts.append(f"{s_err} hata")

        if not parts:
            self._status_var.set("Hiç veri bulunamadı.")
            messagebox.showinfo("Bilgi", "Tarih aralığında okunabilir dosya bulunamadı.")
            return

        self._status_var.set("  |  ".join(parts))
        self._draw_all()

    def _draw_all(self):
        self._draw_chart1()
        self._draw_chart2()
        self._draw_chart3()
        self._draw_chart4()
        self._draw_chart5()
        self._draw_chart6()

    def _on_toggle(self):
        if self._uretim_data or self._sevk_data:
            self._draw_all()

    # ── Veri yardımcıları ─────────────────────────────────────────────────────

    def _u_dates(self):
        return sorted(self._uretim_data)

    def _s_dates(self):
        return sorted(self._sevk_data)

    def _all_dates(self):
        return sorted(set(self._u_dates()) | set(self._s_dates()))

    def _active_fabs(self):
        return [f for f in FACTORIES if self._fab_vars[f].get()]

    def _u_series(self, key, fab, dates):
        return [self._uretim_data.get(d, {}).get(fab, {}).get(key, float("nan"))
                for d in dates]

    def _s_series(self, key, fab, dates):
        return [self._sevk_data.get(d, {}).get(fab, {}).get(key, float("nan"))
                for d in dates]

    def _ratio(self, num_key, fab, dates):
        out = []
        for d in dates:
            fd  = self._uretim_data.get(d, {}).get(fab, {})
            n, t = fd.get(num_key, float("nan")), fd.get("toplam", float("nan"))
            out.append(n / t * 100 if not math.isnan(n) and not math.isnan(t) and t else float("nan"))
        return out

    def _plot_ma(self, ax, x, vals, color, window, label_suffix):
        """MA çizgisi çizer; yeterli veri yoksa sessizce atlar."""
        ma = self._ma(vals, window)
        if any(not math.isnan(v) for v in ma):
            ax.plot(x, ma, color=color, ls=":" if window == 2 else "-.",
                    lw=1.1, alpha=0.65, label=f"{label_suffix} MA{window}")

    # ── Grafik 1: Üretim Trendi ───────────────────────────────────────────────

    def _draw_chart1(self):
        ax = self._axes[0]; ax.clear()
        dates = self._filtered_dates(0, self._u_dates)
        if not dates:
            self._canvases[0].draw(); return
        x = self._to_dt(dates)
        ma2 = self._ma_vars[0][0].get()
        ma6 = self._ma_vars[0][1].get()
        for fab in self._active_fabs():
            for k, (lbl, ls, mk) in C1_METRICS.items():
                if not self._c1_vars[k].get(): continue
                vals = self._u_series(k, fab, dates)
                ax.plot(x, vals, color=FAB_COLORS[fab],
                        ls=ls, marker=mk, markersize=4, lw=1.6,
                        label=f"{fab} – {lbl}")
                tag = f"{fab} – {lbl}"
                if ma2: self._plot_ma(ax, x, vals, FAB_COLORS[fab], 2, tag)
                if ma6: self._plot_ma(ax, x, vals, FAB_COLORS[fab], 6, tag)
        self._fmt_axes(ax, dates)
        ax.set_title("Fabrika Bazında Üretim ve Stok Trendi",
                     fontsize=13, fontweight="bold", pad=10)
        ax.set_ylabel("Miktar (m²)")
        self._legend(ax)
        self._figs[0].tight_layout()
        self._canvases[0].draw()

    # ── Grafik 2: Oran Trendi ─────────────────────────────────────────────────

    def _draw_chart2(self):
        ax = self._axes[1]; ax.clear()
        dates = self._filtered_dates(1, self._u_dates)
        if not dates:
            self._canvases[1].draw(); return
        x = self._to_dt(dates)
        ma2 = self._ma_vars[1][0].get()
        ma6 = self._ma_vars[1][1].get()
        for fab in self._active_fabs():
            for k, (lbl, ls, mk) in C2_RATIOS.items():
                if not self._c2_vars[k].get(): continue
                vals = self._ratio(k.replace("_pct", ""), fab, dates)
                ax.plot(x, vals, color=FAB_COLORS[fab],
                        ls=ls, marker=mk, markersize=4, lw=1.6,
                        label=f"{fab} – {lbl}")
                tag = f"{fab} – {lbl}"
                if ma2: self._plot_ma(ax, x, vals, FAB_COLORS[fab], 2, tag)
                if ma6: self._plot_ma(ax, x, vals, FAB_COLORS[fab], 6, tag)
        self._fmt_axes(ax, dates)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:.1f}%"))
        for ref in (80, 90, 100):
            ax.axhline(ref, color="gray", lw=0.5, ls=":", alpha=0.5)
        ax.set_title("Fabrika Bazında Üretim Oran Trendi",
                     fontsize=13, fontweight="bold", pad=10)
        ax.set_ylabel("Oran (% toplam üretim)")
        self._legend(ax)
        self._figs[1].tight_layout()
        self._canvases[1].draw()

    # ── Grafik 3: Stok Trendi ─────────────────────────────────────────────────

    def _draw_chart3(self):
        if self._ax3r is not None:
            self._ax3r.remove(); self._ax3r = None
        ax = self._axes[2]; ax.clear()
        dates = self._filtered_dates(2, self._u_dates)
        if not dates:
            self._canvases[2].draw(); return

        show_abs   = self._c3_abs_var.get()
        show_pct   = self._c3_pct_var.get()
        show_total = self._c3_total_var.get()
        active     = self._active_fabs()
        ma2        = self._ma_vars[2][0].get()
        ma6        = self._ma_vars[2][1].get()

        all_total = {d: sum((self._uretim_data.get(d, {}).get(f, {}).get("stok") or 0)
                            for f in FACTORIES) for d in dates}
        x = self._to_dt(dates)
        handles, labels = [], []

        ax_pct = ax
        if show_abs and show_pct:
            ax_pct = ax.twinx()
            self._ax3r = ax_pct
            ax_pct.yaxis.set_major_formatter(
                mticker.FuncFormatter(lambda v, _: f"{v:.1f}%"))
            ax_pct.set_ylabel("Oran (% genel toplam stok)")
            ax_pct.spines["top"].set_visible(False)

        for fab in active:
            stok  = self._u_series("stok", fab, dates)
            color = FAB_COLORS[fab]
            if show_abs:
                ln, = ax.plot(x, stok, color=color, ls="-", marker="o",
                              markersize=4, lw=1.8, label=f"{fab} (m²)")
                handles.append(ln); labels.append(f"{fab} (m²)")
                if ma2: self._plot_ma(ax, x, stok, color, 2, f"{fab} (m²)")
                if ma6: self._plot_ma(ax, x, stok, color, 6, f"{fab} (m²)")
            if show_pct:
                pct = [(v / all_total[d] * 100) if all_total[d] and not math.isnan(v)
                       else float("nan") for v, d in zip(stok, dates)]
                ln, = ax_pct.plot(x, pct, color=color, ls="--", marker="s",
                                   markersize=4, lw=1.8, label=f"{fab} (%)")
                handles.append(ln); labels.append(f"{fab} (%)")
                if ma2: self._plot_ma(ax_pct, x, pct, color, 2, f"{fab} (%)")
                if ma6: self._plot_ma(ax_pct, x, pct, color, 6, f"{fab} (%)")

        if show_total:
            total_vals = [all_total[d] for d in dates]
            ln, = ax.plot(x, total_vals, color="black", ls="-", marker="D",
                          markersize=5, lw=2.5, label="Genel Toplam (m²)")
            handles.append(ln); labels.append("Genel Toplam (m²)")
            if ma2: self._plot_ma(ax, x, total_vals, "black", 2, "Genel Toplam (m²)")
            if ma6: self._plot_ma(ax, x, total_vals, "black", 6, "Genel Toplam (m²)")

        self._fmt_axes(ax, dates)
        if show_abs or show_total:
            ax.set_ylabel("Stok (m²)")
        elif show_pct:
            ax.set_ylabel("Oran (%)")
            ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:.1f}%"))
        ax.set_title("Fabrika Stok Trendi", fontsize=13, fontweight="bold", pad=10)
        # MA çizgileri de legend'a eklensin
        extra_h, extra_l = ax.get_legend_handles_labels()
        for h, l in zip(extra_h, extra_l):
            if l not in labels:
                handles.append(h); labels.append(l)
        self._legend(ax, handles, labels)
        self._figs[2].tight_layout()
        self._canvases[2].draw()

    # ── Grafik 4: Sevk Trendi ─────────────────────────────────────────────────

    def _draw_chart4(self):
        ax = self._axes[3]; ax.clear()
        dates = self._filtered_dates(3, self._s_dates)
        if not dates:
            self._canvases[3].draw(); return
        x      = self._to_dt(dates)
        active = self._active_fabs()
        ma2    = self._ma_vars[3][0].get()
        ma6    = self._ma_vars[3][1].get()

        for fab in active:
            for k, (lbl, ls, mk) in SEVK_METRICS.items():
                if not self._c4_vars[k].get(): continue
                vals = self._s_series(k, fab, dates)
                ax.plot(x, vals, color=FAB_COLORS[fab],
                        ls=ls, marker=mk, markersize=4, lw=1.6,
                        label=f"{fab} – {lbl}")
                tag = f"{fab} – {lbl}"
                if ma2: self._plot_ma(ax, x, vals, FAB_COLORS[fab], 2, tag)
                if ma6: self._plot_ma(ax, x, vals, FAB_COLORS[fab], 6, tag)

        if self._c4_total_var.get():
            total = [
                sum(self._sevk_data.get(d, {}).get(f, {}).get("toplam_sevk", 0) or 0
                    for f in FACTORIES)
                for d in dates
            ]
            ax.plot(x, total, color="black", ls="-", marker="D",
                    markersize=5, lw=2.5, label="Günlük Genel Toplam")
            if ma2: self._plot_ma(ax, x, total, "black", 2, "Günlük Genel Toplam")
            if ma6: self._plot_ma(ax, x, total, "black", 6, "Günlük Genel Toplam")

        self._fmt_axes(ax, dates)
        ax.set_title("Fabrika Bazında Sevkiyat Trendi",
                     fontsize=13, fontweight="bold", pad=10)
        ax.set_ylabel("Sevk Miktarı (m²)")
        self._legend(ax)
        self._figs[3].tight_layout()
        self._canvases[3].draw()

    # ── Grafik 5: Üretim vs Sevk ──────────────────────────────────────────────

    def _draw_chart5(self):
        ax = self._axes[4]; ax.clear()
        dates = self._filtered_dates(4, self._all_dates)
        if not dates:
            self._canvases[4].draw(); return
        x      = self._to_dt(dates)
        active = self._active_fabs()
        u_key  = self._c5_uretim_var.get()
        s_key  = self._c5_sevk_var.get()
        u_lbl  = C1_METRICS[u_key][0]
        s_lbl  = SEVK_METRICS[s_key][0]
        ma2    = self._ma_vars[4][0].get()
        ma6    = self._ma_vars[4][1].get()

        for fab in active:
            color  = FAB_COLORS[fab]
            u_vals = [self._uretim_data.get(d, {}).get(fab, {}).get(u_key, float("nan"))
                      for d in dates]
            s_vals = [self._sevk_data.get(d, {}).get(fab, {}).get(s_key, float("nan"))
                      for d in dates]
            ax.plot(x, u_vals, color=color, ls="-", marker="o",
                    markersize=4, lw=1.8, label=f"{fab} – {u_lbl}")
            ax.plot(x, s_vals, color=color, ls="--", marker="^",
                    markersize=4, lw=1.6, alpha=0.8, label=f"{fab} – {s_lbl}")
            if ma2:
                self._plot_ma(ax, x, u_vals, color, 2, f"{fab} – {u_lbl}")
                self._plot_ma(ax, x, s_vals, color, 2, f"{fab} – {s_lbl}")
            if ma6:
                self._plot_ma(ax, x, u_vals, color, 6, f"{fab} – {u_lbl}")
                self._plot_ma(ax, x, s_vals, color, 6, f"{fab} – {s_lbl}")

        self._fmt_axes(ax, dates)
        ax.set_title(f"Üretim ({u_lbl}) vs Sevkiyat ({s_lbl})",
                     fontsize=13, fontweight="bold", pad=10)
        ax.set_ylabel("Miktar (m²)")
        ax.text(0.01, 0.99, "— Üretim  · · · Sevk",
                transform=ax.transAxes, fontsize=8, va="top", color="#555")
        self._legend(ax)
        self._figs[4].tight_layout()
        self._canvases[4].draw()

    # ── Grafik 6: Stok vs Sevk ────────────────────────────────────────────────

    def _draw_chart6(self):
        ax = self._axes[5]; ax.clear()
        dates = self._filtered_dates(5, self._all_dates)
        if not dates:
            self._canvases[5].draw(); return
        x      = self._to_dt(dates)
        active = self._active_fabs()
        s_key  = self._c6_sevk_var.get()
        s_lbl  = SEVK_METRICS[s_key][0]
        ma2    = self._ma_vars[5][0].get()
        ma6    = self._ma_vars[5][1].get()

        for fab in active:
            color = FAB_COLORS[fab]
            stok  = [self._uretim_data.get(d, {}).get(fab, {}).get("stok", float("nan"))
                     for d in dates]
            sevk  = [self._sevk_data.get(d, {}).get(fab, {}).get(s_key, float("nan"))
                     for d in dates]
            ax.plot(x, stok, color=color, ls="-", marker="o",
                    markersize=4, lw=1.8, label=f"{fab} – Stok")
            ax.plot(x, sevk, color=color, ls="--", marker="^",
                    markersize=4, lw=1.6, alpha=0.8, label=f"{fab} – {s_lbl}")
            if ma2:
                self._plot_ma(ax, x, stok, color, 2, f"{fab} – Stok")
                self._plot_ma(ax, x, sevk, color, 2, f"{fab} – {s_lbl}")
            if ma6:
                self._plot_ma(ax, x, stok, color, 6, f"{fab} – Stok")
                self._plot_ma(ax, x, sevk, color, 6, f"{fab} – {s_lbl}")

        self._fmt_axes(ax, dates)
        ax.set_title(f"Stok vs Sevkiyat ({s_lbl})",
                     fontsize=13, fontweight="bold", pad=10)
        ax.set_ylabel("Miktar (m²)")
        ax.text(0.01, 0.99, "— Stok  · · · Sevk",
                transform=ax.transAxes, fontsize=8, va="top", color="#555")
        self._legend(ax)
        self._figs[5].tight_layout()
        self._canvases[5].draw()


# ── Giriş noktası ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    App().mainloop()
