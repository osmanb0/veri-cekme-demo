"""urun-raporu.xlsx'teki gerçek verilerden Bionluk portfolyo görseli (HTML) üretir.

    uv run python portfolyo_gorseli.py   → portfolyo.html (1400x788, ekran görüntüsü alınır)
"""

import html
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

from cek import AZ_STOK

wb = load_workbook("urun-raporu.xlsx")
urunler = list(wb["Ürünler"].iter_rows(min_row=2, values_only=True))
bilgi = dict(wb["Rapor Bilgisi"].iter_rows(values_only=True))

# Kategori özeti: Excel'deki formüllerin aynısını Python'da hesapla (dosyada formül olarak duruyor).
kat = defaultdict(list)
for baslik, kategori, fiyat, stok, puan, *_ in urunler:
    kat[kategori].append((fiyat, stok, puan))
ozet = sorted(
    ((k, len(v), sum(f for f, _, _ in v) / len(v), sum(s for _, s, _ in v)) for k, v in kat.items()
     if k not in ("Default", "Add a comment")),  # sitenin kendi "kategorisiz" etiketleri
    key=lambda r: -r[1],
)[:7]
en_cok = ozet[0][1]

# Tablo örneği: her satır farklı bir gerçek kategoriden; 3'ü az stoklu olsun ki işaretleme görünsün.
gercek = [u for u in urunler if u[1] not in ("Default", "Add a comment")]
ornek, gorulen = [], set()
for az, adet in ((True, 3), (False, 6)):
    for u in gercek:
        if (u[3] <= AZ_STOK) == az and u[1] not in gorulen and len(u[0]) <= 34:
            ornek.append(u)
            gorulen.add(u[1])
            if sum((x[3] <= AZ_STOK) == az for x in ornek) == adet:
                break
ornek.sort(key=lambda u: u[0])


def kisalt(s: str, n: int) -> str:
    return html.escape(s if len(s) <= n else s[: n - 1] + "…")


satirlar = "\n".join(
    f'<tr><td>{kisalt(b, 34)}</td><td>{html.escape(k)}</td><td class="n">£{f:.2f}</td>'
    f'<td class="n{" az" if s <= AZ_STOK else ""}">{s}</td><td class="n">{"★" * p}</td></tr>'
    for b, k, f, s, p, *_ in ornek
)
cubuklar = "\n".join(
    f'<div class="bar"><span>{html.escape(k)}</span><i style="width:{n / en_cok * 100:.0f}%"></i><b>{n}</b></div>'
    for k, n, _, _ in ozet
)

sayfa = f"""<!DOCTYPE html>
<html lang="tr"><head><meta charset="UTF-8"><title>Veri Çekme Portfolyo</title>
<link rel="stylesheet" href="../portfolio-demos/fonts/fraunces.css"><link rel="stylesheet" href="../portfolio-demos/fonts/manrope.css">
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{background:#000;font-family:'Manrope',sans-serif;color:#EDEAE4}}
.stage{{width:1400px;height:788px;overflow:hidden;position:relative;padding:56px 0 0 64px;
  background:radial-gradient(900px 560px at 85% 15%,rgba(129,201,149,.14),transparent 62%),#0F1115}}
.eyebrow{{color:#81C995;font-weight:800;letter-spacing:.26em;text-transform:uppercase;font-size:16px}}
h1{{font-family:'Fraunces',serif;font-weight:900;font-size:54px;line-height:1.05;margin-top:16px;width:470px}}
h1 em{{font-style:italic;color:#81C995}}
.kpi{{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;width:470px;margin-top:28px}}
.kpi div{{background:#1A1D24;border:1px solid rgba(255,255,255,.08);border-radius:14px;padding:14px 18px}}
.kpi b{{display:block;font-size:30px;font-weight:800;color:#fff}}
.kpi span{{font-size:14px;color:#A9ADB6}}
.term{{position:absolute;left:64px;bottom:40px;width:470px;background:#161920;border:1px solid rgba(255,255,255,.08);
  border-radius:12px;padding:14px 16px;font:500 13px/1.65 Consolas,monospace;color:#9BA1AB}}
.term .ok{{color:#81C995}}
.sheet{{position:absolute;left:580px;top:56px;width:760px;background:#fff;color:#1F2328;border-radius:12px;overflow:hidden;
  box-shadow:0 30px 70px rgba(0,0,0,.55);font-size:14px}}
.tabs{{display:flex;gap:2px;background:#E9ECEF;padding:6px 8px 0}}
.tabs span{{padding:6px 14px;border-radius:6px 6px 0 0;color:#555;font-weight:600;font-size:13px}}
.tabs span.on{{background:#fff;color:#1E7E45}}
table{{border-collapse:collapse;width:100%}}
th{{background:#1E7E45;color:#fff;text-align:left;padding:8px 12px;font-weight:700}}
td{{padding:7px 12px;border-bottom:1px solid #E3E6EA;white-space:nowrap}}
tr:nth-child(even) td{{background:#F6F8F7}}
td.n{{text-align:right;font-variant-numeric:tabular-nums}}
td.az{{background:#F8D7DA!important;color:#9C0006;font-weight:700}}
td:last-child{{color:#E0A800;letter-spacing:1px}}
.chart{{position:absolute;left:580px;bottom:40px;width:760px;background:#1A1D24;border:1px solid rgba(255,255,255,.08);
  border-radius:12px;padding:16px 20px}}
.chart h3{{font-size:15px;font-weight:700;margin-bottom:10px;color:#fff}}
.bar{{display:grid;grid-template-columns:150px 1fr 40px;align-items:center;gap:10px;font-size:13px;margin:5px 0;color:#C9CDD4}}
.bar i{{display:block;height:14px;border-radius:4px;background:#81C995}}
.bar b{{text-align:right;color:#fff}}
</style></head><body>
<div class="stage">
  <div class="eyebrow">Veri Çekme · Excel Raporu</div>
  <h1>1000 ürün,<br><em>30 saniyede</em><br>hazır tablo.</h1>
  <div class="kpi">
    <div><b>{bilgi["Ürün sayısı"]}</b><span>ürün çekildi</span></div>
    <div><b>{bilgi["Kategori sayısı"]}</b><span>kategori özetlendi</span></div>
    <div><b>{bilgi[f"Az stoklu ürün (≤{AZ_STOK})"]}</b><span>az stoklu ürün işaretlendi</span></div>
    <div><b>{bilgi["Çekim süresi (sn)"]:.0f} sn</b><span>toplam çekim süresi</span></div>
  </div>
  <div class="term">
    $ python cek.py<br>
    INFO 50 liste sayfasında {bilgi["Ürün sayısı"]} ürün bulundu<br>
    INFO Ürün detayları: {bilgi["Ürün sayısı"]} / {bilgi["Ürün sayısı"]}<br>
    <span class="ok">INFO {bilgi["Ürün sayısı"]} ürün {bilgi["Çekim süresi (sn)"]:.1f} sn'de çekildi → urun-raporu.xlsx</span>
  </div>
  <div class="sheet">
    <div class="tabs"><span class="on">Ürünler</span><span>Kategori Özeti</span><span>Rapor Bilgisi</span></div>
    <table>
      <tr><th>Ürün</th><th>Kategori</th><th>Fiyat</th><th>Stok</th><th>Puan</th></tr>
      {satirlar}
    </table>
  </div>
  <div class="chart">
    <h3>Kategori özeti · en çok ürün</h3>
    {cubuklar}
  </div>
</div>
</body></html>"""

Path("portfolyo.html").write_text(sayfa, encoding="utf-8")
print("portfolyo.html yazıldı")
