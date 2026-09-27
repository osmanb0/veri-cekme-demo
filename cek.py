"""Bir e-ticaret kataloğundan ürün verisi çekip Excel raporu üretir.

Hedef: books.toscrape.com — veri çekme alıştırması için yapılmış, herkese
açık örnek bir kitapçı. Gerçek bir müşteri işinde aynı yapı (liste sayfaları
→ ürün detayları → temizlenmiş tablo → rapor) başka bir siteye uyarlanır.

Kullanım:
    uv run python cek.py                      # tüm katalog (50 sayfa, ~1000 ürün)
    uv run python cek.py --sayfa 3            # ilk 3 sayfa (hızlı deneme)
    uv run python cek.py --cikti rapor.xlsx
"""

import argparse
import asyncio
import logging
import re
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import httpx
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

TABAN = "https://books.toscrape.com/"
ES_ZAMANLI = 8          # aynı anda en fazla bu kadar istek: siteyi yormamak için
DENEME = 3              # geçici ağ hatalarında tekrar sayısı
PUAN = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

log = logging.getLogger("cek")


@dataclass
class Urun:
    baslik: str
    kategori: str
    fiyat_gbp: float
    stok: int
    puan: int
    upc: str
    url: str


async def getir(istemci: httpx.AsyncClient, url: str) -> str:
    for deneme in range(1, DENEME + 1):
        try:
            yanit = await istemci.get(url)
            yanit.raise_for_status()
            # Site UTF-8 döndürüyor ama başlıkta belirtmiyor; httpx'in tahminine bırakma.
            return yanit.content.decode("utf-8")
        except (httpx.TransportError, httpx.HTTPStatusError) as hata:
            if deneme == DENEME:
                raise
            log.warning("%s alınamadı (%s), %d. kez deneniyor", url, type(hata).__name__, deneme + 1)
            await asyncio.sleep(deneme)
    raise AssertionError("ulaşılamaz")


def liste_sayfasi_oku(html: str, sayfa_url: str) -> tuple[list[str], str | None]:
    """Liste sayfasındaki ürün bağlantılarını ve varsa sonraki sayfayı döndürür."""
    soup = BeautifulSoup(html, "html.parser")
    urunler = [urljoin(sayfa_url, a["href"]) for a in soup.select("article.product_pod h3 a")]
    sonraki = soup.select_one("li.next a")
    return urunler, (urljoin(sayfa_url, sonraki["href"]) if sonraki else None)


def urun_sayfasi_oku(html: str, url: str) -> Urun:
    soup = BeautifulSoup(html, "html.parser")
    tablo = {tr.th.get_text(strip=True): tr.td.get_text(strip=True) for tr in soup.select("table.table tr")}
    stok = re.search(r"\((\d+) available\)", tablo.get("Availability", ""))
    puan_sinifi = soup.select_one("p.star-rating")["class"]
    return Urun(
        baslik=soup.select_one("div.product_main h1").get_text(strip=True),
        kategori=soup.select("ul.breadcrumb li")[2].get_text(strip=True),
        fiyat_gbp=float(tablo["Price (incl. tax)"].lstrip("£Â")),
        stok=int(stok.group(1)) if stok else 0,
        puan=next((PUAN[c] for c in puan_sinifi if c in PUAN), 0),
        upc=tablo["UPC"],
        url=url,
    )


async def katalogu_cek(en_fazla_sayfa: int | None) -> list[Urun]:
    sinir = asyncio.Semaphore(ES_ZAMANLI)
    basliklar = {"User-Agent": "veri-cekme-demo/0.1 (portfolyo ornegi)"}
    async with httpx.AsyncClient(timeout=20, headers=basliklar, follow_redirects=True) as istemci:
        # 1) Liste sayfalarını sırayla gez: sonraki sayfa bağlantısını her sayfa kendisi verir.
        urun_linkleri: list[str] = []
        sayfa_url, sayfa_no = urljoin(TABAN, "catalogue/page-1.html"), 0
        while sayfa_url and (en_fazla_sayfa is None or sayfa_no < en_fazla_sayfa):
            linkler, sayfa_url = liste_sayfasi_oku(await getir(istemci, sayfa_url), sayfa_url)
            urun_linkleri += linkler
            sayfa_no += 1
        log.info("%d liste sayfasında %d ürün bulundu", sayfa_no, len(urun_linkleri))

        # 2) Ürün detaylarını sınırlı eşzamanlılıkla çek.
        biten = 0

        async def detay(url: str) -> Urun | None:
            nonlocal biten
            async with sinir:
                try:
                    return urun_sayfasi_oku(await getir(istemci, url), url)
                except Exception as hata:  # tek ürünün bozuk olması raporu durdurmasın
                    log.error("Ürün okunamadı, atlandı: %s (%s)", url, hata)
                    return None
                finally:
                    biten += 1
                    if biten % 100 == 0 or biten == len(urun_linkleri):
                        log.info("Ürün detayları: %d / %d", biten, len(urun_linkleri))

        sonuc = await asyncio.gather(*(detay(u) for u in urun_linkleri))
    return [u for u in sonuc if u is not None]


# --- Excel ---------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", start_color="1E7E45")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
KALIN = Font(bold=True)
AZ_STOK = 2            # bu adet ve altı "az stok" sayılır (bu katalogda ürünlerin ~%11'i)


def sutun_genislikleri(sayfa, genislikler: list[int]) -> None:
    for i, g in enumerate(genislikler, start=1):
        sayfa.column_dimensions[get_column_letter(i)].width = g


def baslik_satiri(sayfa, degerler: list[str]) -> None:
    sayfa.append(degerler)
    for hucre in sayfa[1]:
        hucre.fill, hucre.font = BASLIK_DOLGU, BASLIK_YAZI
        hucre.alignment = Alignment(vertical="center")
    sayfa.freeze_panes = "A2"


def excel_yaz(urunler: list[Urun], yol: Path, sure_sn: float) -> None:
    kitap = Workbook()

    # Sayfa 1: tüm ürünler, filtreli.
    s = kitap.active
    s.title = "Ürünler"
    baslik_satiri(s, ["Ürün", "Kategori", "Fiyat (£)", "Stok", "Puan (1-5)", "UPC", "Bağlantı"])
    for u in sorted(urunler, key=lambda u: (u.kategori, u.baslik)):
        s.append([u.baslik, u.kategori, u.fiyat_gbp, u.stok, u.puan, u.upc, u.url])
    son = s.max_row
    for satir in s.iter_rows(min_row=2, min_col=3, max_col=3):
        satir[0].number_format = "£#,##0.00"
    s.auto_filter.ref = f"A1:G{son}"
    s.conditional_formatting.add(
        f"D2:D{son}",
        CellIsRule(operator="lessThanOrEqual", formula=[str(AZ_STOK)],
                   fill=PatternFill("solid", start_color="F8D7DA"), font=Font(color="9C0006")),
    )
    sutun_genislikleri(s, [60, 22, 11, 8, 11, 19, 60])

    # Sayfa 2: kategori özeti — formüller Excel'de canlı kalır, veri değişince güncellenir.
    o = kitap.create_sheet("Kategori Özeti")
    baslik_satiri(o, ["Kategori", "Ürün Sayısı", "Ort. Fiyat (£)", "Toplam Stok", "Ort. Puan", f"Az Stoklu (≤{AZ_STOK})"])
    kategoriler = sorted({u.kategori for u in urunler}, key=lambda k: -sum(u.kategori == k for u in urunler))
    aralik = lambda sutun: f"'Ürünler'!${sutun}$2:${sutun}${son}"  # noqa: E731
    for i, k in enumerate(kategoriler, start=2):
        o.append([
            k,
            f'=COUNTIF({aralik("B")},A{i})',
            f'=AVERAGEIF({aralik("B")},A{i},{aralik("C")})',
            f'=SUMIF({aralik("B")},A{i},{aralik("D")})',
            f'=AVERAGEIF({aralik("B")},A{i},{aralik("E")})',
            f'=COUNTIFS({aralik("B")},A{i},{aralik("D")},"<={AZ_STOK}")',
        ])
        o.cell(i, 3).number_format = "£#,##0.00"
        o.cell(i, 5).number_format = "0.0"
    sutun_genislikleri(o, [24, 12, 15, 13, 11, 16])

    grafik = BarChart()
    grafik.type, grafik.style = "bar", 10
    grafik.title = "En çok ürün olan 10 kategori"
    grafik.y_axis.title = "Ürün sayısı"
    ilk10 = min(11, o.max_row)
    grafik.add_data(Reference(o, min_col=2, min_row=1, max_row=ilk10), titles_from_data=True)
    grafik.set_categories(Reference(o, min_col=1, min_row=2, max_row=ilk10))
    grafik.legend = None
    grafik.height, grafik.width = 9, 16
    o.add_chart(grafik, "H2")

    # Sayfa 3: rapor bilgisi.
    b = kitap.create_sheet("Rapor Bilgisi")
    bilgiler = [
        ("Kaynak", TABAN),
        ("Oluşturulma", datetime.now().strftime("%d.%m.%Y %H:%M")),
        ("Ürün sayısı", len(urunler)),
        ("Kategori sayısı", len(kategoriler)),
        (f"Az stoklu ürün (≤{AZ_STOK})", sum(u.stok <= AZ_STOK for u in urunler)),
        ("Toplam stok değeri (£)", round(sum(u.fiyat_gbp * u.stok for u in urunler), 2)),
        ("Çekim süresi (sn)", round(sure_sn, 1)),
    ]
    for anahtar, deger in bilgiler:
        b.append([anahtar, deger])
        b.cell(b.max_row, 1).font = KALIN
    b["B6"].number_format = "£#,##0.00"
    sutun_genislikleri(b, [28, 40])

    kitap.save(yol)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--sayfa", type=int, default=None, help="en fazla kaç liste sayfası (varsayılan: hepsi)")
    ap.add_argument("--cikti", type=Path, default=Path("urun-raporu.xlsx"))
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-7s %(message)s", datefmt="%H:%M:%S")
    logging.getLogger("httpx").setLevel(logging.WARNING)

    baslangic = time.perf_counter()
    urunler = asyncio.run(katalogu_cek(args.sayfa))
    sure = time.perf_counter() - baslangic
    if not urunler:
        raise SystemExit("Hiç ürün okunamadı.")
    excel_yaz(urunler, args.cikti, sure)
    log.info("%d ürün %.1f sn'de çekildi → %s", len(urunler), sure, args.cikti.resolve())


if __name__ == "__main__":
    main()
