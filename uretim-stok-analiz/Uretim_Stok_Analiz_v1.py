#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Üretim ve Stok Analiz Paneli
Günlük Üretim ve Stok Raporu PDF dosyalarından fabrika bazında veri çekip
trend grafikleri oluşturur.
Gereksinimler: pip install pdfplumber matplotlib
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re
import math
import threading
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

# Chart 1 — üretim metrikleri: key → (etiket, çizgi stili, marker)
C1_METRICS = {
    "toplam":  ("Toplam Üretim",  "-",       "o"),
    "kal1":    ("1.Kalite",        "--",      "s"),
    "export":  ("Export (Defolu)", "-.",      "^"),
    "iskarta": ("Iskarta",         ":",       "D"),
    "stok":    ("Toplam Stok",     (0,(4,2)), "x"),
}

# Chart 2 — oran metrikleri
C2_RATIOS = {
    "kal1_pct":    ("1.Kalite %",  "-",  "o"),
    "export_pct":  ("Export %",    "--", "^"),
    "iskarta_pct": ("Iskarta %",   ":",  "D"),
}

DATE_RE = re.compile(r"(\d{2})-(\d{2})-(\d{4})")
_TARGET = set(FACTORIES)

rcParams["font.family"] = "sans-serif"

# ── PDF Ayrıştırıcı ───────────────────────────────────────────────────────────

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


def extract_pdf(path: Path) -> Optional[dict]:
    """
    Döndürür: {fabrika: {kal1, export, iskarta, toplam, stok}} veya None
    """
    try:
        with pdfplumber.open(str(path)) as pdf:
            result = {}

            # Sayfa 1 → Tablo 1 → Günlük Üretim
            tables_p1 = pdf.pages[0].extract_tables()
            if tables_p1:
                for row in tables_p1[0][2:]:   # ilk 2 satır başlık
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

            # Sayfa 3 → Tablo 1 → Fabrika Stok
            if len(pdf.pages) > 2:
                tables_p3 = pdf.pages[2].extract_tables()
                if tables_p3:
                    for row in tables_p3[0][1:]:   # ilk satır başlık
                        if not row or len(row) < 3:
                            continue
                        fab = (row[1] or "").strip()
                        if fab in _TARGET:
                            result.setdefault(fab, {
                                "kal1": float("nan"), "export": float("nan"),
                                "iskarta": float("nan"), "toplam": float("nan"),
                            })["stok"] = _num(row[2])

            return result or None
    except Exception as exc:
        print(f"  ✗ {path.name}: {exc}")
        return None


def date_from_path(path: Path) -> Optional[date]:
    m = DATE_RE.search(path.stem)
    if not m:
        return None
    try:
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


# ── Yardımcı widget: DD.MM.YYYY giriş ────────────────────────────────────────

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
        self.minsize(1200, 720)
        self.resizable(True, True)

        # Veri deposu: {date → {fabrika → {kal1, export, iskarta, toplam, stok}}}
        self._data: dict[date, dict] = {}
        self._files: list[Path] = []

        # Fabrika toggle'ları (tüm grafikleri etkiler)
        self._fab_vars: dict[str, tk.BooleanVar] = {
            f: tk.BooleanVar(value=True) for f in FACTORIES
        }

        # Grafik 1 metrik toggle'ları
        self._c1_vars: dict[str, tk.BooleanVar] = {
            k: tk.BooleanVar(value=(k in {"toplam", "kal1"}))
            for k in C1_METRICS
        }

        # Grafik 2 oran toggle'ları
        self._c2_vars: dict[str, tk.BooleanVar] = {
            k: tk.BooleanVar(value=True) for k in C2_RATIOS
        }

        # Grafik 3 toggle'ları
        self._c3_abs_var   = tk.BooleanVar(value=True)   # miktarsal
        self._c3_pct_var   = tk.BooleanVar(value=False)  # oran %
        self._c3_total_var = tk.BooleanVar(value=True)   # genel toplam

        # Grafik figürleri ve canvas'ları (build_ui'de oluşturulur)
        self._fig1 = self._ax1 = self._canvas1 = None
        self._fig2 = self._ax2 = self._canvas2 = None
        self._fig3 = self._ax3 = self._ax3r = self._canvas3 = None

        self._style_setup()
        self._build_ui()

    # ── Stil ─────────────────────────────────────────────────────────────────

    def _style_setup(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame",      background=self.BG)
        s.configure("TLabel",      background=self.BG, font=("Helvetica", 11))
        s.configure("Bold.TLabel", background=self.BG,
                    font=("Helvetica", 12, "bold"))
        s.configure("TLabelframe", background=self.CARD, relief="solid",
                    borderwidth=1, bordercolor="#D1D9E0")
        s.configure("TLabelframe.Label", background=self.CARD,
                    foreground="#1A73E8", font=("Helvetica", 11, "bold"))
        s.configure("TCheckbutton", background=self.CARD, font=("Helvetica", 10))
        s.configure("TButton",      font=("Helvetica", 11), padding=(8, 5))
        s.configure("Big.TButton",  font=("Helvetica", 12, "bold"), padding=(10, 8))
        s.configure("TEntry",       fieldbackground="white", font=("Helvetica", 11))
        s.configure("TNotebook",    background=self.BG)
        s.configure("TNotebook.Tab", font=("Helvetica", 11), padding=(12, 6))

    # ── UI İnşası ────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Sol panel + sağ notebook yan yana
        left = ttk.Frame(self, padding=(12, 12, 6, 12))
        left.pack(side="left", fill="y")

        right = ttk.Frame(self, padding=(6, 12, 12, 12))
        right.pack(side="left", fill="both", expand=True)

        self._build_left(left)
        self._build_notebook(right)

        # Alt durum çubuğu
        self._status_var = tk.StringVar(value="Dosya yüklenmedi.")
        ttk.Label(self, textvariable=self._status_var,
                  foreground="#718096", background=self.BG,
                  font=("Helvetica", 10)).pack(side="bottom", anchor="w",
                                               padx=12, pady=4)

    # ── Sol Panel ────────────────────────────────────────────────────────────

    def _build_left(self, parent):
        parent.configure(width=230)

        # Başlık
        ttk.Label(parent, text="Analiz Paneli",
                  font=("Helvetica", 14, "bold"),
                  foreground="#1A73E8",
                  background=self.BG).pack(anchor="w", pady=(0, 12))

        # ── Dosya seçimi
        ff = ttk.LabelFrame(parent, text="  Dosyalar  ", padding=(8, 6))
        ff.pack(fill="x", pady=(0, 10))

        self._lb = tk.Listbox(ff, height=7, selectmode="extended",
                              font=("Helvetica", 9), bg="white",
                              relief="flat", borderwidth=1)
        self._lb.pack(fill="x")

        btn_row = ttk.Frame(ff)
        btn_row.pack(fill="x", pady=(4, 0))
        ttk.Button(btn_row, text="+ Ekle",  command=self._add_files).pack(side="left", padx=(0, 4))
        ttk.Button(btn_row, text="Sil",     command=self._remove_file).pack(side="left", padx=(0, 4))
        ttk.Button(btn_row, text="Temizle", command=self._clear_files).pack(side="left")

        # ── Tarih aralığı filtresi
        df = ttk.LabelFrame(parent, text="  Tarih Aralığı  ", padding=(8, 6))
        df.pack(fill="x", pady=(0, 10))

        ttk.Label(df, text="Başlangıç:", background=self.CARD,
                  font=("Helvetica", 10)).grid(row=0, column=0, sticky="w")
        self._start_e = DateEntry(df, initial=date.today().replace(day=1))
        self._start_e.configure(background=self.CARD)
        self._start_e.grid(row=0, column=1, sticky="w", pady=2)

        ttk.Label(df, text="Bitiş:", background=self.CARD,
                  font=("Helvetica", 10)).grid(row=1, column=0, sticky="w")
        self._end_e = DateEntry(df, initial=date.today())
        self._end_e.configure(background=self.CARD)
        self._end_e.grid(row=1, column=1, sticky="w", pady=2)

        # ── Fabrika toggle'ları
        fabf = ttk.LabelFrame(parent, text="  Fabrikalar  ", padding=(8, 6))
        fabf.pack(fill="x", pady=(0, 10))

        for fab in FACTORIES:
            cb_frame = ttk.Frame(fabf)
            cb_frame.pack(anchor="w", pady=1)
            # Renkli kare göstergesi
            dot = tk.Canvas(cb_frame, width=12, height=12, bg=self.CARD,
                            highlightthickness=0)
            dot.create_rectangle(1, 1, 11, 11,
                                  fill=FAB_COLORS[fab], outline="")
            dot.pack(side="left", padx=(0, 4))
            ttk.Checkbutton(
                cb_frame, text=fab,
                variable=self._fab_vars[fab],
                command=self._on_toggle,
            ).pack(side="left")

        # ── Grafikle düğmesi
        ttk.Button(parent, text="▶  Grafikle",
                   command=self._load_and_plot,
                   style="Big.TButton").pack(fill="x", pady=(4, 0))

    # ── Notebook (3 grafik sekmesi) ──────────────────────────────────────────

    def _build_notebook(self, parent):
        nb = ttk.Notebook(parent)
        nb.pack(fill="both", expand=True)
        self._nb = nb

        # Sekme 1 — Üretim Trendi
        t1 = ttk.Frame(nb)
        nb.add(t1, text="📊  Üretim Trendi")
        self._build_tab1(t1)

        # Sekme 2 — Oran Trendi
        t2 = ttk.Frame(nb)
        nb.add(t2, text="📈  Oran Trendi")
        self._build_tab2(t2)

        # Sekme 3 — Stok Trendi
        t3 = ttk.Frame(nb)
        nb.add(t3, text="🏭  Stok Trendi")
        self._build_tab3(t3)

    def _build_tab1(self, parent):
        ctrl = tk.Frame(parent, bg=self.BG, pady=4)
        ctrl.pack(fill="x", padx=8)
        tk.Label(ctrl, text="Göster:", bg=self.BG,
                 font=("Helvetica", 10, "bold")).pack(side="left", padx=(0, 8))
        for key, (label, ls, mk) in C1_METRICS.items():
            ttk.Checkbutton(ctrl, text=label,
                            variable=self._c1_vars[key],
                            command=self._draw_chart1).pack(side="left", padx=4)

        self._fig1 = Figure(figsize=(10, 5.5), dpi=100, facecolor=self.CARD)
        self._ax1  = self._fig1.add_subplot(111)
        self._canvas1 = FigureCanvasTkAgg(self._fig1, parent)
        self._canvas1.get_tk_widget().pack(fill="both", expand=True)
        NavigationToolbar2Tk(self._canvas1, parent)

    def _build_tab2(self, parent):
        ctrl = tk.Frame(parent, bg=self.BG, pady=4)
        ctrl.pack(fill="x", padx=8)
        tk.Label(ctrl, text="Göster:", bg=self.BG,
                 font=("Helvetica", 10, "bold")).pack(side="left", padx=(0, 8))
        for key, (label, ls, mk) in C2_RATIOS.items():
            ttk.Checkbutton(ctrl, text=label,
                            variable=self._c2_vars[key],
                            command=self._draw_chart2).pack(side="left", padx=4)

        self._fig2 = Figure(figsize=(10, 5.5), dpi=100, facecolor=self.CARD)
        self._ax2  = self._fig2.add_subplot(111)
        self._canvas2 = FigureCanvasTkAgg(self._fig2, parent)
        self._canvas2.get_tk_widget().pack(fill="both", expand=True)
        NavigationToolbar2Tk(self._canvas2, parent)

    def _build_tab3(self, parent):
        ctrl = tk.Frame(parent, bg=self.BG, pady=4)
        ctrl.pack(fill="x", padx=8)
        tk.Label(ctrl, text="Göster:", bg=self.BG,
                 font=("Helvetica", 10, "bold")).pack(side="left", padx=(0, 8))
        ttk.Checkbutton(ctrl, text="Miktarsal (m²)",
                        variable=self._c3_abs_var,
                        command=self._draw_chart3).pack(side="left", padx=4)
        ttk.Checkbutton(ctrl, text="Oran (% toplam stok)",
                        variable=self._c3_pct_var,
                        command=self._draw_chart3).pack(side="left", padx=4)
        ttk.Separator(ctrl, orient="vertical").pack(side="left", fill="y",
                                                     padx=8, pady=2)
        ttk.Checkbutton(ctrl, text="Genel Toplam Stok",
                        variable=self._c3_total_var,
                        command=self._draw_chart3).pack(side="left", padx=4)

        self._fig3 = Figure(figsize=(10, 5.5), dpi=100, facecolor=self.CARD)
        self._ax3  = self._fig3.add_subplot(111)
        self._ax3r = None  # ikincil eksen (oran için)
        self._canvas3 = FigureCanvasTkAgg(self._fig3, parent)
        self._canvas3.get_tk_widget().pack(fill="both", expand=True)
        NavigationToolbar2Tk(self._canvas3, parent)

    # ── Dosya Yönetimi ────────────────────────────────────────────────────────

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Üretim ve Stok Raporu PDF'lerini Seçin",
            filetypes=[("PDF", "*.pdf"), ("Tümü", "*.*")],
        )
        for p in paths:
            pp = Path(p)
            if pp not in self._files:
                self._files.append(pp)
                self._lb.insert("end", pp.name)

    def _remove_file(self):
        for i in reversed(self._lb.curselection()):
            self._lb.delete(i)
            self._files.pop(i)

    def _clear_files(self):
        self._lb.delete(0, "end")
        self._files.clear()

    # ── Veri Yükleme ─────────────────────────────────────────────────────────

    def _load_and_plot(self):
        if not self._files:
            messagebox.showwarning("Uyarı", "Lütfen önce PDF dosyası ekleyin.")
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
            data: dict[date, dict] = {}
            ok = skip = err = 0
            for path in self._files:
                d = date_from_path(path)
                if d is None or not (d_start <= d <= d_end):
                    skip += 1
                    continue
                result = extract_pdf(path)
                if result:
                    data[d] = result
                    ok += 1
                else:
                    err += 1
            self.after(0, self._on_loaded, data, ok, skip, err)

        threading.Thread(target=worker, daemon=True).start()

    def _on_loaded(self, data, ok, skip, err):
        self._data = data
        n = len(data)
        if n == 0:
            self._status_var.set("Hiç veri bulunamadı.")
            messagebox.showinfo("Bilgi", "Tarih aralığında okunabilir dosya bulunamadı.")
            return

        dates = sorted(data.keys())
        self._status_var.set(
            f"{ok} dosya yüklendi | {skip} atlandı | {err} hata  "
            f"({dates[0].strftime('%d.%m.%Y')} – {dates[-1].strftime('%d.%m.%Y')})"
        )
        self._draw_all()

    def _draw_all(self):
        self._draw_chart1()
        self._draw_chart2()
        self._draw_chart3()

    # ── Toggle tepkisi ────────────────────────────────────────────────────────

    def _on_toggle(self):
        if self._data:
            self._draw_all()

    # ── Yardımcı seriler ──────────────────────────────────────────────────────

    def _dates(self) -> list[date]:
        return sorted(self._data.keys())

    def _active_fabs(self) -> list[str]:
        return [f for f in FACTORIES if self._fab_vars[f].get()]

    def _series(self, key: str, fab: str, dates: list) -> list:
        return [self._data.get(d, {}).get(fab, {}).get(key, float("nan"))
                for d in dates]

    def _ratio(self, num_key: str, fab: str, dates: list) -> list:
        out = []
        for d in dates:
            fd = self._data.get(d, {}).get(fab, {})
            n = fd.get(num_key, float("nan"))
            t = fd.get("toplam", float("nan"))
            if math.isnan(n) or math.isnan(t) or t == 0:
                out.append(float("nan"))
            else:
                out.append(n / t * 100)
        return out

    @staticmethod
    def _to_dt(dates: list) -> list:
        return [datetime(d.year, d.month, d.day) for d in dates]

    @staticmethod
    def _fmt_axes(ax, dates: list):
        """Tarih formatı ve bin ayracı."""
        n = len(dates)
        if n <= 31:
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m"))
        elif n <= 90:
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m.%y"))
        else:
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%m.%Y"))
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        ax.figure.autofmt_xdate(rotation=35, ha="right")
        ax.yaxis.set_major_formatter(
            mticker.FuncFormatter(lambda v, _: f"{int(v):,}".replace(",", "."))
        )
        ax.grid(True, alpha=0.25, linestyle="--")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

    @staticmethod
    def _legend_outside(ax, handles=None, labels=None):
        """Efsaneyi grafiğin sağına koy."""
        if handles is None:
            handles, labels = ax.get_legend_handles_labels()
        if handles:
            ax.legend(handles, labels,
                      loc="upper left",
                      bbox_to_anchor=(1.01, 1),
                      borderaxespad=0,
                      fontsize=9,
                      framealpha=0.9)

    # ── Grafik 1: Üretim Trendi ───────────────────────────────────────────────

    def _draw_chart1(self):
        ax = self._ax1
        ax.clear()

        dates = self._dates()
        if not dates or not self._data:
            self._canvas1.draw()
            return

        x = self._to_dt(dates)
        active_fabs = self._active_fabs()

        for fab in active_fabs:
            color = FAB_COLORS[fab]
            for key, (label, ls, mk) in C1_METRICS.items():
                if not self._c1_vars[key].get():
                    continue
                vals = self._series(key, fab, dates)
                ax.plot(x, vals, color=color, ls=ls, marker=mk,
                        markersize=4, linewidth=1.6,
                        label=f"{fab} – {label}")

        self._fmt_axes(ax, dates)
        ax.set_title("Fabrika Bazında Üretim ve Stok Trendi",
                     fontsize=13, fontweight="bold", pad=10)
        ax.set_ylabel("Miktar (m²)")
        self._legend_outside(ax)
        self._fig1.tight_layout()
        self._canvas1.draw()

    # ── Grafik 2: Oran Trendi ─────────────────────────────────────────────────

    def _draw_chart2(self):
        ax = self._ax2
        ax.clear()

        dates = self._dates()
        if not dates or not self._data:
            self._canvas2.draw()
            return

        x = self._to_dt(dates)
        active_fabs = self._active_fabs()

        for fab in active_fabs:
            color = FAB_COLORS[fab]
            for key, (label, ls, mk) in C2_RATIOS.items():
                if not self._c2_vars[key].get():
                    continue
                num_key = key.replace("_pct", "")   # kal1, export, iskarta
                vals = self._ratio(num_key, fab, dates)
                ax.plot(x, vals, color=color, ls=ls, marker=mk,
                        markersize=4, linewidth=1.6,
                        label=f"{fab} – {label}")

        self._fmt_axes(ax, dates)
        ax.set_title("Fabrika Bazında Üretim Oran Trendi",
                     fontsize=13, fontweight="bold", pad=10)
        ax.set_ylabel("Oran (% toplam üretim)")
        ax.yaxis.set_major_formatter(
            mticker.FuncFormatter(lambda v, _: f"{v:.1f}%")
        )
        # Referans çizgileri
        for ref in (80, 90, 100):
            ax.axhline(ref, color="gray", linewidth=0.5, linestyle=":", alpha=0.5)
        self._legend_outside(ax)
        self._fig2.tight_layout()
        self._canvas2.draw()

    # ── Grafik 3: Stok Trendi ─────────────────────────────────────────────────

    def _draw_chart3(self):
        # İkincil ekseni sıfırla
        if self._ax3r is not None:
            self._ax3r.remove()
            self._ax3r = None
        ax = self._ax3
        ax.clear()

        dates = self._dates()
        if not dates or not self._data:
            self._canvas3.draw()
            return

        show_abs   = self._c3_abs_var.get()
        show_pct   = self._c3_pct_var.get()
        show_total = self._c3_total_var.get()
        active_fabs = self._active_fabs()

        # Her gün için 5 fabrikanın toplam stoğu (oran paydasında kullanılacak)
        all_total = {
            d: sum(
                (self._data.get(d, {}).get(f, {}).get("stok") or 0)
                for f in FACTORIES
            )
            for d in dates
        }

        x = self._to_dt(dates)
        handles, labels = [], []

        # İkincil eksen sadece her iki gösterim aynı anda aktifse
        if show_abs and show_pct:
            ax_pct = ax.twinx()
            self._ax3r = ax_pct
            ax_pct.yaxis.set_major_formatter(
                mticker.FuncFormatter(lambda v, _: f"{v:.1f}%")
            )
            ax_pct.set_ylabel("Oran (% genel toplam stok)", fontsize=10)
            ax_pct.spines["top"].set_visible(False)
        else:
            ax_pct = ax

        for fab in active_fabs:
            color = FAB_COLORS[fab]
            stok = self._series("stok", fab, dates)

            if show_abs:
                ln, = ax.plot(x, stok, color=color, ls="-", marker="o",
                              markersize=4, linewidth=1.8,
                              label=f"{fab} (m²)")
                handles.append(ln); labels.append(f"{fab} (m²)")

            if show_pct:
                pct = [
                    (v / all_total[d] * 100)
                    if all_total[d] and not math.isnan(v) else float("nan")
                    for v, d in zip(stok, dates)
                ]
                ln, = ax_pct.plot(x, pct, color=color, ls="--", marker="s",
                                   markersize=4, linewidth=1.8,
                                   label=f"{fab} (%)")
                handles.append(ln); labels.append(f"{fab} (%)")

        if show_total:
            total_vals = [all_total[d] for d in dates]
            ln, = ax.plot(x, total_vals, color="black", ls="-", marker="D",
                          markersize=5, linewidth=2.5,
                          label="Genel Toplam (m²)")
            handles.append(ln); labels.append("Genel Toplam (m²)")

        self._fmt_axes(ax, dates)

        if show_abs or show_total:
            ax.set_ylabel("Stok (m²)")
        elif show_pct:
            ax.set_ylabel("Oran (%)")
            ax.yaxis.set_major_formatter(
                mticker.FuncFormatter(lambda v, _: f"{v:.1f}%")
            )

        ax.set_title("Fabrika Stok Trendi", fontsize=13, fontweight="bold", pad=10)
        self._legend_outside(ax, handles, labels)
        self._fig3.tight_layout()
        self._canvas3.draw()


# ── Giriş noktası ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    App().mainloop()
