# -*- coding: utf-8 -*-
"""
Günlük Detaylı Sevk Raporu PDF'lerinden UNIQUE ÜNVAN bazında toplam m² hesaplama uygulaması.

Kullanım:
1) pip install pdfplumber pandas openpyxl
2) python unvan_m2_toplam_uygulamasi.py
3) PDF dosyasını seçin.
4) Program aynı klasöre Excel çıktısı oluşturur.

Not: Sadece miktar birimi "M2" olan satırları dikkate alır. ADET satırları hariç tutulur.
"""

import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from collections import defaultdict

import pdfplumber
import pandas as pd


# Satır yapısı genellikle: ... UNVAN EBAT SERI KAL URUN ADI Miktar M2 Fiyat Ort.Fiyat Tutar
# Bu uygulama önce M2 miktarını bulur, sonra miktarın sol tarafındaki müşteri ünvanını ebat başlangıcına kadar ayrıştırır.
EBAT_PATTERN = re.compile(r"\b\d{2,3}(?:[Xx,\.\-]\d{2,3})?(?:[Xx,\.\-]?\d{1,2})?N?\b|\bSTANDA\b|\b\d+\s*CM\b", re.IGNORECASE)
M2_PATTERN = re.compile(r"(?P<miktar>\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})\s+M2\b", re.IGNORECASE)

# Raporlarda bu alanlardan sonra ünvan başlar. Bölge müdürlüğü bazen 1-3 kelime sürebiliyor.
PREFIX_CODES = {"KOS", "BOL", "PRO", "IHR", "FAB"}


def tr_number_to_float(value: str) -> float:
    """Türkçe sayı formatını float'a çevirir: 1.234,56 -> 1234.56"""
    return float(value.replace(".", "").replace(",", "."))


def float_to_tr(value: float) -> str:
    """Float sayıyı Türkçe formatlar: 1234.56 -> 1.234,56"""
    return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def clean_unvan(unvan: str) -> str:
    unvan = re.sub(r"\s+", " ", unvan).strip()
    return unvan


def extract_unvan_from_left_part(left_part: str) -> str | None:
    """
    M2 miktarının solundaki metinden ünvanı tahmin eder.
    Örnek sol taraf:
    KOS BATI KARADENİZ İMAMOĞULLARI İNŞ... 61X61 TORTUGA 1.K 61X61...
    """
    parts = left_part.split()
    if not parts:
        return None

    # İlk kodu atla: KOS/BOL/PRO/IHR/FAB
    if parts[0] in PREFIX_CODES:
        rest = " ".join(parts[1:])
    else:
        rest = left_part

    # Ünvan sonunu ilk ebat/STANDA/CM başlangıcından önce kabul et.
    ebat_match = EBAT_PATTERN.search(rest)
    if not ebat_match:
        return None

    before_ebat = rest[:ebat_match.start()].strip()
    words = before_ebat.split()

    # Bölge müdürlüğü kısmını atmak için pratik kural:
    # Raporda genellikle ilk 1-3 kelime bölge, sonrasında ünvan gelir.
    known_region_starts = [
        "DIYARBAKIR DOĞU", "BATI KARADENİZ", "GÜNEY DOĞU ANA", "DOĞU KARADENİZ",
        "İÇANADOLU", "İÇ ANADOLU", "EGE", "MARMARA", "YURTDIŞI SATIŞLA",
        "BULGARISTAN", "GRUP FİRMA", "PROJE SATIŞLAR-S", "PROJE SATIŞLAR-B",
        "PROJE SATIŞLAR-G", "PROJE SATIŞLAR-A", "ADAPAZARI SATIŞL", "YOK"
    ]

    upper = before_ebat.upper()
    for region in known_region_starts:
        if upper.startswith(region):
            candidate = before_ebat[len(region):].strip()
            return clean_unvan(candidate) if candidate else None

    # Fallback: ilk iki kelimeyi bölge varsay, kalan ünvan.
    if len(words) > 2:
        return clean_unvan(" ".join(words[2:]))
    return clean_unvan(before_ebat)


def parse_pdf(pdf_path: str) -> pd.DataFrame:
    records = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            for raw_line in text.splitlines():
                line = re.sub(r"\s+", " ", raw_line).strip()
                if not line or " Toplam " in f" {line} ":
                    continue

                m2_match = M2_PATTERN.search(line)
                if not m2_match:
                    continue

                miktar_text = m2_match.group("miktar")
                miktar = tr_number_to_float(miktar_text)
                left_part = line[:m2_match.start()].strip()
                unvan = extract_unvan_from_left_part(left_part)

                if unvan:
                    records.append({
                        "Sayfa": page_no,
                        "Ünvan": unvan,
                        "Miktar (m²)": miktar,
                        "Satır": line
                    })

    if not records:
        return pd.DataFrame(columns=["Ünvan", "Toplam Miktar (m²)", "Geçtiği Satır Sayısı"])

    detail_df = pd.DataFrame(records)
    summary_df = (
        detail_df.groupby("Ünvan", as_index=False)
        .agg(**{
            "Toplam Miktar (m²)": ("Miktar (m²)", "sum"),
            "Geçtiği Satır Sayısı": ("Miktar (m²)", "count")
        })
        .sort_values("Toplam Miktar (m²)", ascending=False)
    )
    return summary_df, detail_df


def run_app():
    root = tk.Tk()
    root.withdraw()

    pdf_path = filedialog.askopenfilename(
        title="Günlük Detaylı Sevk Raporu PDF seçin",
        filetypes=[("PDF Dosyaları", "*.pdf")]
    )
    if not pdf_path:
        return

    try:
        summary_df, detail_df = parse_pdf(pdf_path)
        base = os.path.splitext(pdf_path)[0]
        output_path = base + "_unvan_m2_toplam.xlsx"

        # Görsel format için ayrıca Türkçe formatlı kolon eklenir.
        summary_export = summary_df.copy()
        summary_export["Toplam Miktar (m²) - Formatlı"] = summary_export["Toplam Miktar (m²)"].apply(float_to_tr)

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            summary_export.to_excel(writer, sheet_name="Unvan Toplamları", index=False)
            detail_df.to_excel(writer, sheet_name="Satır Detayı", index=False)

        messagebox.showinfo(
            "İşlem tamamlandı",
            f"Ünvan bazlı toplamlar oluşturuldu:\n{output_path}"
        )
    except Exception as exc:
        messagebox.showerror("Hata", f"İşlem sırasında hata oluştu:\n{exc}")


if __name__ == "__main__":
    run_app()
