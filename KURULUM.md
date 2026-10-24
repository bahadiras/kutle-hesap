# Kütle Hesap — Kurulum

**Uygulama adresi:** https://bahadiras.github.io/kutle-hesap/

## Klasördeki dosyalar
| Dosya / klasör | Ne işe yarar |
|---|---|
| `KutleHesap.html` | Tek dosyalık yedek. Bilgisayarda çift tıkla, telefonda Chrome ile aç. İnternet gerekmez. |
| `docs/` | GitHub Pages'te yayınlanan PWA dosyaları (otomatik üretilir, elle düzenleme). |
| `src/` | Uygulamanın kaynak dosyaları (arayüz, service worker, manifest). |
| `build.py` | Excel'i okuyup `docs/` ve `KutleHesap.html` dosyalarını yeniden üretir. |
| `profil_bilgileri.xlsx` | Profil verisi. |

## 1. Telefona kurma (bir kez, internet gerekir)
1. Yukarıdaki adresi telefonda **Chrome** ile aç.
2. **⋮ → Uygulamayı yükle** (veya **Ana ekrana ekle**).
3. Ana ekrandaki **Kütle Hesap** ikonuyla aç. Artık internet olmadan çalışır.
4. Menüden (☰) **Yedek uygulama dosyasını indir**. `KutleHesap.html` İndirilenler klasörüne kaydedilir. Uygulama bir gün silinirse bu dosyayı Chrome ile açarak internetsiz kullanmaya devam edersin.

Bilgisayarda: aynı adresi Chrome veya Edge ile aç, adres çubuğundaki **yükle** ikonuna tıkla.

## 2. Hesap listesini korumak
- Menü → **Listeyi yedekle**: liste bir `.json` dosyasına kaydedilir.
- Menü → **Yedekten geri yükle**: bu dosyayı geri yükler (telefon ↔ bilgisayar arasında taşımak için de kullanılabilir).

## 3. Güncelleme yayınlama
1. `profil_bilgileri.xlsx` veya `src/` içindekileri düzenle.
   Excel'de sütunların sırası serbesttir; sütunlar başlık adına göre bulunur (`G [kg/m]` mutlaka olmalı).
   Bir sütun eksik ya da bir değer hatalıysa `build.py` hatayı yazar ve uygulamayı üretmez.
2. Bu klasörde terminalde:
   ```
   python build.py
   git add -A
   git commit -m "Güncelleme"
   git push
   ```
3. 1–2 dakika içinde site güncellenir. Telefondaki uygulama, internete bağlıyken bir kez açıldığında kendini günceller ("Uygulama güncellendi → Yenile").
