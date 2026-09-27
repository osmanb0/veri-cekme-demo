# Veri Çekme → Excel Raporu (portfolyo demosu)

Bir e-ticaret kataloğunun tüm ürünlerini gezip ürün adı, kategori, fiyat, stok
ve puan bilgisini toplayan; sonucu filtreli, biçimlendirilmiş ve özet sayfalı
bir Excel raporuna döken Python betiği.

Hedef site [books.toscrape.com](https://books.toscrape.com/), veri çekme
alıştırması için yapılmış herkese açık örnek bir kitapçıdır. Müşteri işinde
aynı akış (liste sayfaları → ürün detayları → temiz tablo → rapor) başka bir
siteye uyarlanır.

## Çalıştırma

```powershell
uv sync
uv run python cek.py              # tüm katalog: 50 sayfa, 1000 ürün (~30 sn)
uv run python cek.py --sayfa 3    # hızlı deneme
uv run python cek.py --cikti rapor.xlsx
```

## Rapor içeriği (`urun-raporu.xlsx`)

- **Ürünler:** 1000 satır, otomatik filtre, dondurulmuş başlık, `£` biçimi;
  stoğu 2 ve altındaki ürünler kırmızı işaretli.
- **Kategori Özeti:** kategori başına ürün sayısı, ortalama fiyat, toplam
  stok, ortalama puan, az stoklu ürün sayısı. Değerler Excel formülüdür
  (`COUNTIF`, `AVERAGEIF`, `SUMIF`, `COUNTIFS`), veri değişince güncellenir.
  En kalabalık 10 kategori için çubuk grafik.
- **Rapor Bilgisi:** kaynak, tarih, adetler, toplam stok değeri, çekim süresi.

## Tasarım notları

- En fazla 8 eşzamanlı istek: siteyi yormadan hızlı çalışır.
- Geçici ağ hatalarında 3 kez yeniden dener; tek bir bozuk ürün sayfası
  raporu durdurmaz, loglanıp atlanır.
- `portfolyo_gorseli.py`, rapordaki gerçek verilerden portfolyo görselini üretir.
