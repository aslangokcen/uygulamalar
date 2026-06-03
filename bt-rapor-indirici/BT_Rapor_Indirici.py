#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BT Mail Rapor İndirici
BT Reporting mail eklerini seçilen tarih aralığında otomatik bulup indirir.
Mac'te Mail.app, Windows'ta Outlook kullanır.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import os
import re
import shutil
import platform
import subprocess
import threading
from datetime import date, datetime, timedelta
from pathlib import Path

SYSTEM = platform.system()  # 'Darwin', 'Windows', 'Linux'
SENDER_KEYWORD = "BT Reporting"

REPORTS = [
    {
        "id": "gunluk_sevk",
        "name": "Günlük Sevk Raporu",
        "keyword": "Günlük Sevk Raporu",
        "file_pattern": "Gunluk Sevk Raporu",
    },
    {
        "id": "ic_piyasa",
        "name": "İç Piyasa Sevk Raporu (Yıllık Karşılaştırmalı)",
        "keyword": "İç Piyasa Sevk",
        "file_pattern": "Ic Piyasa Sevk Raporu",
    },
    {
        "id": "ihracat",
        "name": "İhracat Sevk Raporu (Yıllık Karşılaştırmalı)",
        "keyword": "İhracat Sevk",
        "file_pattern": "Ihracat Sevk Raporu",
    },
    {
        "id": "gecikmis",
        "name": "Gecikmiş Sevkler",
        "keyword": "Gecikmiş Sevkler",
        "file_pattern": "Gecikmis Sevkler",
    },
    {
        "id": "fiyat_miktar",
        "name": "Fiyat ve Miktarsal Olarak Eşleşmeyen Malzeme Kabulleri",
        "keyword": "Fiyat ve Miktarsal",
        "file_pattern": "Fiyat ve Miktarsal",
    },
    {
        "id": "uretim_urun",
        "name": "Günlük Üretim Raporu - Ürün Bazlı",
        "keyword": "Günlük Üretim Raporu",
        "file_pattern": "Gunluk Uretim Raporu",
    },
    {
        "id": "fabrika_stok",
        "name": "Fabrika Stok ve Rezervleri Raporu",
        "keyword": "Fabrika Stok",
        "file_pattern": "Fabrika Stok ve Rezervleri Raporu",
    },
    {
        "id": "uretim_stok",
        "name": "Üretim ve Stok Raporu",
        "keyword": "Üretim ve Stok",
        "file_pattern": "Uretim ve Stok Raporu",
    },
    {
        "id": "detayli_sevk",
        "name": "Günlük Detaylı Sevk Raporu",
        "keyword": "Günlük Detaylı Sevk",
        "file_pattern": "Gunluk Detayli Sevk Raporu",
    },
    {
        "id": "detayli_satin",
        "name": "Günlük Detaylı Satın Alma Raporu",
        "keyword": "Günlük Detaylı Satın",
        "file_pattern": "Gunluk Detayli Satin Alma Raporu",
    },
    {
        "id": "sevk_karsilastirma",
        "name": "Sevk Karşılaştırma Raporu",
        "keyword": "Sevk Karşılaştırma",
        "file_pattern": "Sevk Karsilastirma Raporu",
    },
    {
        "id": "ik_fazla_mesai",
        "name": "İnsan Kaynakları - Fazla Mesai Raporu",
        "keyword": "Fazla Mesai",
        "file_pattern": "Insan Kaynaklari",
    },
]

ATTACHMENT_EXTENSIONS = {".pdf", ".xlsx", ".xls", ".csv"}


# ─────────────────────────────────────────────────────────────────────────────
# File organizer helpers
# ─────────────────────────────────────────────────────────────────────────────

_DATE_RE = re.compile(r"(\d{2})-(\d{2})-(\d{4})")

_TR_MONTHS = {
    "01": "Ocak",    "02": "Şubat",   "03": "Mart",   "04": "Nisan",
    "05": "Mayıs",   "06": "Haziran", "07": "Temmuz", "08": "Ağustos",
    "09": "Eylül",   "10": "Ekim",    "11": "Kasım",  "12": "Aralık",
}

# Sorted longest-first so specific patterns match before generic ones
_PATTERNS = sorted(REPORTS, key=lambda r: len(r["file_pattern"]), reverse=True)


def _parse_file(p: Path):
    """Return (month_label, report_label) extracted from the filename."""
    stem = p.stem
    m = _DATE_RE.search(stem)
    if m:
        dd, mm, yyyy = m.group(1), m.group(2), m.group(3)
        month_label = f"{yyyy}-{mm} {_TR_MONTHS.get(mm, mm)}"
    else:
        month_label = "Tarihi Belirsiz"

    report_label = "Diğer"
    for r in _PATTERNS:
        if r["file_pattern"].lower() in stem.lower():
            report_label = r["name"]
            break

    return month_label, report_label


def _collect_files(dest: Path):
    """Recursively collect all matching files under dest."""
    return [p for p in dest.rglob("*")
            if p.is_file() and p.suffix.lower() in ATTACHMENT_EXTENSIONS]


def _remove_empty_dirs(root: Path):
    """Remove empty subdirectories bottom-up, never root itself."""
    for d in sorted(root.rglob("*"), key=lambda p: len(p.parts), reverse=True):
        if d.is_dir() and d != root:
            try:
                d.rmdir()
            except OSError:
                pass


def organize_files(dest: Path, mode: str):
    """
    Move files under dest into a subfolder structure.
    mode: 'month' | 'name' | 'month_name'
    Returns (moved, skipped, errors, log_lines).
    """
    files = _collect_files(dest)
    moved = skipped = errors = 0
    log_lines = []

    for p in files:
        month_label, report_label = _parse_file(p)

        if mode == "month":
            target_dir = dest / month_label
        elif mode == "name":
            target_dir = dest / report_label
        else:  # month_name
            target_dir = dest / month_label / report_label

        if p.parent == target_dir:
            skipped += 1
            log_lines.append(("skip", f"Zaten yerinde: {p.name}"))
            continue

        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / p.name

        if target.exists():
            skipped += 1
            log_lines.append(("skip", f"Hedefte mevcut, atlandı: {p.name}"))
        else:
            try:
                shutil.move(str(p), str(target))
                moved += 1
                log_lines.append(("saved", f"→ {target.relative_to(dest)}"))
            except Exception as exc:
                errors += 1
                log_lines.append(("error", f"Hata ({p.name}): {exc}"))

    _remove_empty_dirs(dest)
    return moved, skipped, errors, log_lines


# ─────────────────────────────────────────────────────────────────────────────
# AppleScript generation (macOS)
# ─────────────────────────────────────────────────────────────────────────────

MONTH_EN = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December",
}


def _as_date_block(var_name: str, d: date) -> str:
    """Build locale-safe AppleScript date via component assignment."""
    return (
        f"set {var_name} to current date\n"
        f"set year of {var_name} to {d.year}\n"
        f"set month of {var_name} to {d.month}\n"
        f"set day of {var_name} to {d.day}\n"
        f"set time of {var_name} to 0\n"
    )


def build_applescript(selected_reports, start_date, end_date,
                      dest_folder: Path, all_mailboxes: bool) -> str:
    dest = str(dest_folder).replace('"', '\\"')

    # Build AppleScript list of keywords  {"kw1", "kw2", ...}
    kw_list = "{" + ", ".join(f'"{r["keyword"]}"' for r in selected_reports) + "}"

    end_exclusive = end_date + timedelta(days=1)

    mailbox_block = ""
    if all_mailboxes:
        mailbox_block = """
        set theMailboxes to mailboxes of anAccount
        repeat with aMailbox in theMailboxes"""
        mailbox_end = "        end repeat"
    else:
        mailbox_block = """
        try
            set aMailbox to mailbox "INBOX" of anAccount
        on error
            try
                set aMailbox to mailbox "Gelen Kutusu" of anAccount
            on error
                set aMailbox to mailbox 1 of anAccount
            end try
        end try
        set theMailboxes to {aMailbox}
        repeat with aMailbox in theMailboxes"""
        mailbox_end = "        end repeat"

    script = f"""
-- BT Mail Rapor İndirici AppleScript
set destFolder to "{dest}"
set theKeywords to {kw_list}
set savedCount to 0
set skippedCount to 0
set logLines to {{}}

{_as_date_block("startDate", start_date)}{_as_date_block("endDate", end_exclusive)}
-- Create destination folder
do shell script "mkdir -p " & quoted form of destFolder

tell application "Mail"
    set theAccounts to accounts
    repeat with anAccount in theAccounts
        set acctName to name of anAccount
        set end of logLines to "ACCOUNT:" & acctName
        {mailbox_block}
            try
                set theMessages to (every message of aMailbox whose ¬
                    sender contains "{SENDER_KEYWORD}" and ¬
                    date received >= startDate and ¬
                    date received < endDate)
                repeat with aMessage in theMessages
                    set msgSubject to subject of aMessage
                    set kwMatched to false
                    repeat with kw in theKeywords
                        if msgSubject contains kw then
                            set kwMatched to true
                            exit repeat
                        end if
                    end repeat
                    if kwMatched then
                        try
                            set theAttachments to mail attachments of aMessage
                            repeat with anAtt in theAttachments
                                set attName to name of anAtt
                                set attLower to do shell script "echo " & quoted form of attName & " | tr '[:upper:]' '[:lower:]'"
                                if attLower ends with ".pdf" or attLower ends with ".xlsx" or attLower ends with ".xls" or attLower ends with ".csv" then
                                    set savePath to destFolder & "/" & attName
                                    -- Skip if file already exists
                                    try
                                        set testAlias to (POSIX file savePath) as alias
                                        set skippedCount to skippedCount + 1
                                        set end of logLines to "SKIP:" & attName
                                    on error
                                        -- File doesn't exist, save it
                                        try
                                            save anAtt in POSIX file savePath
                                            set savedCount to savedCount + 1
                                            set end of logLines to "SAVED:" & attName
                                        on error errMsg
                                            set end of logLines to "ERROR:" & attName & " - " & errMsg
                                        end try
                                    end try
                                end if
                            end repeat
                        on error
                            -- Could not access attachments
                        end try
                    end if
                end repeat
            on error
                -- Could not access mailbox
            end try
        {mailbox_end}
    end repeat
end tell

-- Return structured log
set logOutput to ""
repeat with aLine in logLines
    set logOutput to logOutput & aLine & linefeed
end repeat
set logOutput to logOutput & "TOTAL_SAVED:" & savedCount & linefeed
set logOutput to logOutput & "TOTAL_SKIPPED:" & skippedCount & linefeed
return logOutput
"""
    return script


# ─────────────────────────────────────────────────────────────────────────────
# Simple date entry widget (DD.MM.YYYY)
# ─────────────────────────────────────────────────────────────────────────────

class DateEntry(tk.Frame):
    def __init__(self, parent, initial_date: date = None, **kw):
        super().__init__(parent, **kw)
        d = initial_date or date.today()
        self.day_var   = tk.StringVar(value=f"{d.day:02d}")
        self.month_var = tk.StringVar(value=f"{d.month:02d}")
        self.year_var  = tk.StringVar(value=str(d.year))

        vcmd = (self.register(lambda v: v == "" or v.isdigit()), "%P")
        for var, w in [(self.day_var, 3), (self.month_var, 3), (self.year_var, 5)]:
            e = ttk.Entry(self, textvariable=var, width=w, justify="center",
                          validate="key", validatecommand=vcmd)
            e.pack(side="left")
            if w != 5:
                ttk.Label(self, text=".").pack(side="left")

    def get(self) -> date:
        try:
            return date(int(self.year_var.get()),
                        int(self.month_var.get()),
                        int(self.day_var.get()))
        except (ValueError, TypeError) as e:
            raise ValueError(f"Geçersiz tarih: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# Main application
# ─────────────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    COLORS = {
        "bg":     "#F0F4F8",
        "accent": "#1A73E8",
        "panel":  "#FFFFFF",
        "border": "#D1D9E0",
        "text":   "#2D3748",
        "muted":  "#718096",
        "green":  "#38A169",
        "red":    "#E53E3E",
        "yellow": "#D69E2E",
    }

    def __init__(self):
        super().__init__()
        self.title("BT Mail Rapor İndirici")
        self.configure(bg=self.COLORS["bg"])
        self.resizable(True, True)
        self.minsize(680, 760)
        self._check_vars: dict[str, tk.BooleanVar] = {}
        self._running = False
        self._style_setup()
        self._build_ui()

    # ── Styling ──────────────────────────────────────────────────────────────

    def _style_setup(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame",       background=self.COLORS["bg"])
        s.configure("Card.TFrame",  background=self.COLORS["panel"],
                    relief="flat", borderwidth=1)
        s.configure("TLabel",       background=self.COLORS["bg"],
                    foreground=self.COLORS["text"], font=("Helvetica", 11))
        s.configure("Card.TLabel",  background=self.COLORS["panel"],
                    foreground=self.COLORS["text"], font=("Helvetica", 11))
        s.configure("Title.TLabel", background=self.COLORS["bg"],
                    foreground=self.COLORS["accent"],
                    font=("Helvetica", 18, "bold"))
        s.configure("Section.TLabel", background=self.COLORS["panel"],
                    foreground=self.COLORS["text"],
                    font=("Helvetica", 12, "bold"))
        s.configure("TCheckbutton",  background=self.COLORS["panel"],
                    foreground=self.COLORS["text"], font=("Helvetica", 11))
        s.configure("Accent.TButton", font=("Helvetica", 12, "bold"),
                    padding=(16, 8))
        s.configure("TButton", font=("Helvetica", 11), padding=(10, 6))
        s.configure("TLabelframe",   background=self.COLORS["panel"],
                    relief="solid", borderwidth=1,
                    bordercolor=self.COLORS["border"])
        s.configure("TLabelframe.Label", background=self.COLORS["panel"],
                    foreground=self.COLORS["accent"],
                    font=("Helvetica", 11, "bold"))
        s.configure("TEntry",  fieldbackground="white",
                    font=("Helvetica", 11))
        s.configure("Horizontal.TProgressbar",
                    troughcolor=self.COLORS["border"],
                    background=self.COLORS["accent"], thickness=6)

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        wrap = ttk.Frame(self, padding=20)
        wrap.pack(fill="both", expand=True)

        # Title
        ttk.Label(wrap, text="BT Mail Rapor İndirici",
                  style="Title.TLabel").pack(anchor="w", pady=(0, 4))
        ttk.Label(wrap, text="Mail uygulamasından BT Reporting eklerini otomatik indirir.",
                  foreground=self.COLORS["muted"], background=self.COLORS["bg"],
                  font=("Helvetica", 11)).pack(anchor="w", pady=(0, 16))

        # ── Report types ──────────────────────────────────────────────────────
        rep_frame = ttk.LabelFrame(wrap, text="  Rapor Türleri  ", padding=(12, 8))
        rep_frame.pack(fill="x", pady=(0, 12))

        header_row = ttk.Frame(rep_frame)
        header_row.pack(fill="x", pady=(0, 6))

        self._select_all_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            header_row, text="Tümünü Seç / Kaldır",
            variable=self._select_all_var,
            command=self._toggle_all,
            style="TCheckbutton",
        ).pack(side="left")

        sep = ttk.Separator(rep_frame, orient="horizontal")
        sep.pack(fill="x", pady=(0, 6))

        # Two-column grid for checkboxes
        grid = ttk.Frame(rep_frame)
        grid.pack(fill="x")
        for i, report in enumerate(REPORTS):
            var = tk.BooleanVar(value=True)
            self._check_vars[report["id"]] = var
            col, row = i % 2, i // 2
            ttk.Checkbutton(
                grid, text=report["name"], variable=var,
                command=self._update_select_all,
            ).grid(row=row, column=col, sticky="w", padx=(0, 20), pady=2)
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)

        # ── Date range ────────────────────────────────────────────────────────
        date_frame = ttk.LabelFrame(wrap, text="  Tarih Aralığı  ", padding=(12, 8))
        date_frame.pack(fill="x", pady=(0, 12))

        date_row = ttk.Frame(date_frame)
        date_row.pack(fill="x")
        today = date.today()
        week_ago = today - timedelta(days=7)

        ttk.Label(date_row, text="Başlangıç:", style="Card.TLabel").pack(side="left", padx=(0, 6))
        self._start_entry = DateEntry(date_row, initial_date=week_ago)
        self._start_entry.pack(side="left", padx=(0, 24))

        ttk.Label(date_row, text="Bitiş:", style="Card.TLabel").pack(side="left", padx=(0, 6))
        self._end_entry = DateEntry(date_row, initial_date=today)
        self._end_entry.pack(side="left")

        # ── Options ───────────────────────────────────────────────────────────
        opt_frame = ttk.LabelFrame(wrap, text="  Seçenekler  ", padding=(12, 8))
        opt_frame.pack(fill="x", pady=(0, 12))

        self._all_mailboxes_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opt_frame,
            text="Tüm klasörleri tara (yavaş, ama kapsamlı — önerilen: BT mailleri Gelen Kutusu'ndaysa işaretlemeyin)",
            variable=self._all_mailboxes_var,
        ).pack(anchor="w")

        # ── Destination folder ────────────────────────────────────────────────
        dest_frame = ttk.LabelFrame(wrap, text="  Kayıt Klasörü  ", padding=(12, 8))
        dest_frame.pack(fill="x", pady=(0, 12))

        dest_row = ttk.Frame(dest_frame)
        dest_row.pack(fill="x")

        default_dest = str(Path.home() / "Desktop" / "BT Mailleri")
        self._dest_var = tk.StringVar(value=default_dest)
        ttk.Entry(dest_row, textvariable=self._dest_var,
                  font=("Helvetica", 11)).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(dest_row, text="Seç…", command=self._choose_folder).pack(side="left")

        # ── Action buttons ────────────────────────────────────────────────────
        btn_row = ttk.Frame(wrap)
        btn_row.pack(fill="x", pady=(0, 8))

        self._start_btn = ttk.Button(
            btn_row, text="▶  İndirmeyi Başlat",
            command=self._start_download, style="Accent.TButton",
        )
        self._start_btn.pack(side="left", padx=(0, 10))

        ttk.Button(btn_row, text="📁  Klasörü Aç",
                   command=self._open_dest).pack(side="left")

        self._status_lbl = ttk.Label(
            btn_row, text="", foreground=self.COLORS["muted"],
            background=self.COLORS["bg"], font=("Helvetica", 10),
        )
        self._status_lbl.pack(side="right")

        # ── Organize buttons ──────────────────────────────────────────────────
        org_frame = ttk.LabelFrame(wrap, text="  Dosyaları Düzenle  ", padding=(12, 8))
        org_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(
            org_frame,
            text="İndirilen klasördeki dosyaları seçilen düzene göre alt klasörlere taşır.",
            foreground=self.COLORS["muted"],
            background=self.COLORS["panel"],
            font=("Helvetica", 10),
        ).pack(anchor="w", pady=(0, 8))

        org_btn_row = ttk.Frame(org_frame)
        org_btn_row.pack(fill="x")

        ttk.Button(
            org_btn_row, text="📅  Aya Göre",
            command=lambda: self._organize("month"),
        ).pack(side="left", padx=(0, 8))

        ttk.Button(
            org_btn_row, text="📂  Rapora Göre",
            command=lambda: self._organize("name"),
        ).pack(side="left", padx=(0, 8))

        ttk.Button(
            org_btn_row, text="📅 → 📂  Ay / Rapor",
            command=lambda: self._organize("month_name"),
        ).pack(side="left")

        # ── Progress bar ──────────────────────────────────────────────────────
        self._progress = ttk.Progressbar(wrap, mode="indeterminate",
                                          style="Horizontal.TProgressbar")
        self._progress.pack(fill="x", pady=(0, 8))

        # ── Log ───────────────────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(wrap, text="  İşlem Günlüğü  ", padding=(8, 6))
        log_frame.pack(fill="both", expand=True)

        self._log_box = scrolledtext.ScrolledText(
            log_frame, height=12, state="disabled",
            font=("Courier New", 10) if SYSTEM == "Windows" else ("Menlo", 10),
            bg="#1E1E2E", fg="#CDD6F4", insertbackground="white",
            relief="flat",
        )
        self._log_box.pack(fill="both", expand=True)
        # Tag colors for log lines
        self._log_box.tag_config("saved",   foreground="#A6E3A1")  # green
        self._log_box.tag_config("skip",    foreground="#F9E2AF")  # yellow
        self._log_box.tag_config("error",   foreground="#F38BA8")  # red
        self._log_box.tag_config("account", foreground="#89DCEB")  # cyan
        self._log_box.tag_config("info",    foreground="#CDD6F4")  # default
        self._log_box.tag_config("summary", foreground="#CBA6F7", font=(
            "Menlo" if SYSTEM == "Darwin" else "Courier New", 10, "bold"))

    # ── Checkbox helpers ──────────────────────────────────────────────────────

    def _toggle_all(self):
        v = self._select_all_var.get()
        for var in self._check_vars.values():
            var.set(v)

    def _update_select_all(self):
        self._select_all_var.set(all(v.get() for v in self._check_vars.values()))

    # ── Folder helpers ────────────────────────────────────────────────────────

    def _choose_folder(self):
        d = filedialog.askdirectory(initialdir=self._dest_var.get())
        if d:
            self._dest_var.set(d)

    def _open_dest(self):
        dest = Path(self._dest_var.get())
        if not dest.exists():
            messagebox.showinfo("Bilgi", "Klasör henüz oluşturulmadı.")
            return
        if SYSTEM == "Darwin":
            subprocess.run(["open", str(dest)])
        elif SYSTEM == "Windows":
            subprocess.run(["explorer", str(dest)])
        else:
            subprocess.run(["xdg-open", str(dest)])

    # ── Download orchestration ────────────────────────────────────────────────

    def _start_download(self):
        if self._running:
            return

        selected = [r for r in REPORTS if self._check_vars[r["id"]].get()]
        if not selected:
            messagebox.showwarning("Uyarı", "Lütfen en az bir rapor türü seçin.")
            return

        try:
            start = self._start_entry.get()
            end   = self._end_entry.get()
        except ValueError as exc:
            messagebox.showerror("Tarih Hatası", str(exc))
            return

        if start > end:
            messagebox.showerror("Tarih Hatası",
                                  "Başlangıç tarihi bitiş tarihinden sonra olamaz.")
            return

        dest = Path(self._dest_var.get())

        self._running = True
        self._start_btn.configure(state="disabled")
        self._progress.start(12)
        self._status_lbl.configure(text="Aranıyor…")

        self._log("─" * 64, "info")
        self._log(f"Başlatıldı: {datetime.now().strftime('%H:%M:%S')}", "info")
        self._log(f"Tarih aralığı: {start.strftime('%d.%m.%Y')} — {end.strftime('%d.%m.%Y')}", "info")
        self._log(f"Seçilen rapor: {len(selected)}", "info")
        self._log(f"Hedef: {dest}", "info")
        self._log("─" * 64, "info")

        threading.Thread(
            target=self._worker,
            args=(selected, start, end, dest),
            daemon=True,
        ).start()

    def _worker(self, selected, start, end, dest):
        try:
            dest.mkdir(parents=True, exist_ok=True)
            if SYSTEM == "Darwin":
                self._run_macos(selected, start, end, dest)
            elif SYSTEM == "Windows":
                self._run_windows(selected, start, end, dest)
            else:
                self._log("⚠ Desteklenmeyen işletim sistemi.", "error")
        except Exception as exc:
            self._log(f"Beklenmeyen hata: {exc}", "error")
        finally:
            self.after(0, self._done)

    # ── macOS (Mail.app via AppleScript) ──────────────────────────────────────

    def _run_macos(self, selected, start, end, dest):
        self._log("Mail.app'e bağlanılıyor…", "info")
        script = build_applescript(
            selected, start, end, dest,
            all_mailboxes=self._all_mailboxes_var.get(),
        )
        result = subprocess.run(
            ["osascript", "-"],
            input=script,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        if result.returncode != 0:
            err = result.stderr.strip()
            self._log(f"AppleScript hatası: {err}", "error")
            if "Not authorized" in err or "permission" in err.lower():
                self._log(
                    "Mail.app erişim izni gerekli. Sistem Ayarları → Gizlilik → "
                    "Otomasyon bölümünden Terminal'e Mail.app iznini verin.",
                    "error",
                )
            return

        # Parse structured output
        saved_count = 0
        skipped_count = 0
        for line in result.stdout.splitlines():
            if line.startswith("ACCOUNT:"):
                self._log(f"  Hesap: {line[8:]}", "account")
            elif line.startswith("SAVED:"):
                fname = line[6:]
                self._log(f"  ✓ {fname}", "saved")
                saved_count += 1
            elif line.startswith("SKIP:"):
                fname = line[5:]
                self._log(f"  ↷ Atlandı (zaten var): {fname}", "skip")
                skipped_count += 1
            elif line.startswith("ERROR:"):
                self._log(f"  ✗ {line[6:]}", "error")
            elif line.startswith("TOTAL_SAVED:"):
                saved_count = int(line.split(":")[1])
            elif line.startswith("TOTAL_SKIPPED:"):
                skipped_count = int(line.split(":")[1])

        self._log("─" * 64, "summary")
        self._log(
            f"Tamamlandı — İndirilen: {saved_count}  |  Atlanan: {skipped_count}",
            "summary",
        )

    # ── Windows (Outlook via win32com) ────────────────────────────────────────

    def _run_windows(self, selected, start, end, dest):
        try:
            import win32com.client  # type: ignore
        except ImportError:
            self._log(
                "Windows için 'pywin32' paketi gerekli.\n"
                "Kurmak için: pip install pywin32",
                "error",
            )
            return

        self._log("Outlook'a bağlanılıyor…", "info")
        try:
            outlook   = win32com.client.Dispatch("Outlook.Application")
            ns        = outlook.GetNamespace("MAPI")
        except Exception as exc:
            self._log(f"Outlook başlatılamadı: {exc}", "error")
            return

        saved_count   = 0
        skipped_count = 0

        # Build Outlook restriction filter (EN date format MM/DD/YYYY)
        start_str = start.strftime("%m/%d/%Y 00:00 AM")
        end_str   = (end + timedelta(days=1)).strftime("%m/%d/%Y 00:00 AM")
        restrict  = (
            f"[SenderName] like '%{SENDER_KEYWORD}%' AND "
            f"[ReceivedTime] >= '{start_str}' AND "
            f"[ReceivedTime] < '{end_str}'"
        )

        def search_folder(folder):
            nonlocal saved_count, skipped_count
            try:
                items = folder.Items.Restrict(restrict)
                for msg in items:
                    subject = getattr(msg, "Subject", "") or ""
                    for report in selected:
                        if report["keyword"] in subject:
                            if msg.Attachments.Count > 0:
                                for att in msg.Attachments:
                                    name = att.FileName
                                    ext = Path(name).suffix.lower()
                                    if ext in ATTACHMENT_EXTENSIONS:
                                        save_path = dest / name
                                        if save_path.exists():
                                            self._log(f"  ↷ Atlandı: {name}", "skip")
                                            skipped_count += 1
                                        else:
                                            att.SaveAsFile(str(save_path))
                                            self._log(f"  ✓ {name}", "saved")
                                            saved_count += 1
                            break
            except Exception as exc:
                self._log(f"  Klasör hatası: {exc}", "error")

            if self._all_mailboxes_var.get():
                for sf in folder.Folders:
                    search_folder(sf)

        for account in ns.Accounts:
            self._log(f"  Hesap: {account.DisplayName}", "account")
            try:
                inbox = ns.GetDefaultFolder(6)  # olFolderInbox
                search_folder(inbox)
            except Exception as exc:
                self._log(f"    Gelen Kutusu hatası: {exc}", "error")

        self._log("─" * 64, "summary")
        self._log(
            f"Tamamlandı — İndirilen: {saved_count}  |  Atlanan: {skipped_count}",
            "summary",
        )

    # ── File organizer ────────────────────────────────────────────────────────

    def _organize(self, mode: str):
        if self._running:
            messagebox.showwarning("Uyarı", "İndirme devam ederken düzenleme yapılamaz.")
            return

        dest = Path(self._dest_var.get())
        if not dest.exists():
            messagebox.showinfo("Bilgi", "Klasör henüz oluşturulmadı.")
            return

        files = _collect_files(dest)
        if not files:
            messagebox.showinfo("Bilgi", "Klasörde düzenlenecek dosya bulunamadı.")
            return

        labels = {
            "month":      "Aya Göre",
            "name":       "Rapora Göre",
            "month_name": "Ay → Rapor",
        }
        if not messagebox.askyesno(
            "Onay",
            f"{len(files)} dosya '{labels[mode]}' düzenine göre taşınacak.\nDevam edilsin mi?",
        ):
            return

        self._running = True
        self._start_btn.configure(state="disabled")
        self._progress.start(12)
        self._status_lbl.configure(text="Düzenleniyor…")
        self._log("─" * 64, "info")
        self._log(f"Düzenleme başladı: {labels[mode]}  ({len(files)} dosya)", "info")

        def worker():
            try:
                moved, skipped, errors, lines = organize_files(dest, mode)
                for tag, msg in lines:
                    self._log(f"  {msg}", tag)
                self._log("─" * 64, "summary")
                self._log(
                    f"Tamamlandı — Taşınan: {moved}  |  Atlanan: {skipped}  |  Hata: {errors}",
                    "summary",
                )
            except Exception as exc:
                self._log(f"Beklenmeyen hata: {exc}", "error")
            finally:
                self.after(0, self._done)

        threading.Thread(target=worker, daemon=True).start()

    # ── UI helpers ────────────────────────────────────────────────────────────

    def _done(self):
        self._progress.stop()
        self._start_btn.configure(state="normal")
        self._status_lbl.configure(text="Hazır")
        self._running = False

    def _log(self, message: str, tag: str = "info"):
        self.after(0, self._append_log, message, tag)

    def _append_log(self, message: str, tag: str):
        self._log_box.configure(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self._log_box.insert("end", f"[{ts}] {message}\n", tag)
        self._log_box.see("end")
        self._log_box.configure(state="disabled")


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    App().mainloop()
