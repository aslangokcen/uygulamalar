# Uygulamalar

Üretim, raporlama ve finans alanında günlük işleri kolaylaştıran küçük, bağımsız masaüstü ve tarayıcı araçları koleksiyonu. Her araç tek başına çalışır; ortak bir kurulum gerektirmez.

> Genel amaçlı, kuruma özgü olmayan araçlar. Örnek/anonim verilerle çalışır.

## İçindekiler

| Klasör | Araç | Tür | Çalıştırma |
|--------|------|-----|------------|
| [`bt-rapor-indirici/`](bt-rapor-indirici/) | **Mail Rapor İndirici** — mail eklerini (PDF/Excel) tarih aralığına göre otomatik bulup indirir ve aya/türe göre klasörler. macOS Mail.app + Windows Outlook. | Python (Tkinter) | `python3 BT_Rapor_Indirici.py` |
| [`uretim-stok-analiz/`](uretim-stok-analiz/) | **Üretim & Stok Analiz Paneli** — PDF raporlardan fabrika bazlı üretim/stok/sevk trend grafikleri. v1 → v2 → v3 (gün filtresi + hareketli ortalama). | Python (pdfplumber, matplotlib) | `python3 Uretim_Stok_Analiz_v3.py` |
| [`unvan-m2-toplam/`](unvan-m2-toplam/) | **Ünvan Bazlı m² Toplamı** — sevk PDF'inden müşteri ünvanı bazında toplam m² hesaplar, Excel çıktısı verir. | Python (pdfplumber, pandas, openpyxl) | `python3 unvan_m2_toplam_uygulamasi.py` |
| [`sarf-analiz-masasi/`](sarf-analiz-masasi/) | **Sarf Analiz Masası** — Excel dosyalarını yükleyip ürün/fabrika/yıl kıyasları ve mutabakat üretir. | React / Vite | `npm install && npm run dev` |
| [`mevduat-fiyat-vade/`](mevduat-fiyat-vade/) | **Teklif Kıyaslama** — vadeli satın alma tekliflerini mevduat faizine göre bugünkü değere indirgeyip karşılaştırır. | HTML | Tarayıcıda aç |

## Kullanım

- **HTML araçları:** `.html` dosyasını çift tıklayıp tarayıcıda açın.
- **Python araçları:** `python3 dosya_adi.py`. Gerekli paketler:
  ```bash
  pip install pdfplumber matplotlib pandas openpyxl
  ```

## Lisans

[MIT](LICENSE) — özgürce kullanabilir, değiştirebilir ve dağıtabilirsiniz.
