# HAMU Tools V1

**HAMU · Hüseyin Avni UZUN**

Microsoft Excel için veri düzenleme, tablo, tarih, sayfa, denetim ve aktarım araçları. Ribbon ve Dosya > HAMU Tools üzerinden erişilir. Türkçe ve İngilizce arayüzü destekler.

## İndir ve kur

1. [HAMUToolsV1.xlam dosyasını indir](https://github.com/huseyinavniuzun/HAMU-Tools/raw/refs/heads/main/08_Addin/HAMUToolsV1.xlam).
2. Dosyayı kalıcı bir klasöre yerleştir. İnternetten gelen dosyada Windows güvenlik engeli varsa dosya özelliklerinden yalnızca güvendiğin bu dosyanın engelini kaldır.
3. Excel'de **Dosya > Seçenekler > Eklentiler > Yönet: Excel Eklentileri > Git > Gözat** yolundan dosyayı seç.
4. Eski HAMU sürümünü devre dışı bırakıp Excel'i yeniden aç.

Excel'in Windows masaüstü sürümünü ve VBA desteğini gerektirir. Karekod/barkod üretiminde kullanılan çevrimiçi servisler internet bağlantısı gerektirebilir. VBA veri işlemleri Excel'in standart geri alma geçmişini etkileyebilir; önemli dosyalarında önce yedek oluştur.

## Neler var?

- Veri birleştirme, ayıklama, bölme, gruplama, pivot ve unpivot.
- Metin temizleme, Türkçe harf dönüşümü ve Normal İfade ile bul/değiştir.
- Hızlı tablo, veri girişi, sütun seçimi ve tablo düzenleme.
- Tarih seçici, tarih düzeltme ve tarih kontrolü.
- Sayfa indeksi, toplu fotoğraf, barkod/karekod ve çıktı araçları.
- Veri kalitesi, vurgulama, sayım ve HAMU hücre fonksiyonları.
- Yedekleme, klasör içeriği listeleme ve aktarma sihirbazı.

## Kaynak dosyaları

| Klasör | İçerik |
|---|---|
| `01_CustomUI` | Güncel eklentiden çıkarılmış Ribbon XML, ilişkiler ve gömülü ikonlar |
| `02_Modules` | Core ve Features VBA modülleri |
| `03_Class_Modules` | VBA sınıfları |
| `04_ThisWorkbook` | Çalışma kitabı/sayfa nesnesi kodları |
| `05_Forms` | UserForm `.frm` ve eşleşen `.frx` dosyaları |
| `06_Icons` | İkon kaynakları ve boyut çeşitleri |
| `07_Documentation` | Kullanım, geliştirme, metin düzenleme ve performans belgeleri |
| `08_Addin` | Kullanıma hazır `.xlam` eklentisi |
| `09_Templates` | LAMBDA şablonu |

VBA kaynakları güncel `.xlam` içinden dışa aktarılmıştır. Git, `.gitattributes` sayesinde VBA metinlerini depoda UTF-8 olarak saklar, klonlanan çalışma klasöründe Windows-1254/CRLF olarak üretir. Excel'e kaynak aktaracaksan `git clone` kullan; GitHub'dan tek tek Raw metin indirmek UTF-8 verebilir. `.frm` ve `.frx` çiftlerini birlikte tut. ThisWorkbook kodunu mevcut kitap nesnesine yerleştir; sıradan sınıf olarak içe aktarma.

Dosyadaki bir modülü değiştirmek hazır `.xlam` dosyasını kendiliğinden güncellemez. Excel VBA düzenleyicisinde içe aktar, derle ve eklentiyi kaydet. Ribbon düzenlemeleri için Office RibbonX Editor kullanılabilir.

## Belgeler

- [Kullanım rehberi](07_Documentation/Kullanim_Rehberi.md)
- [Geliştirme rehberi](07_Documentation/Gelistirme_Rehberi.md)
- [Hücre fonksiyonları](07_Documentation/HAMU_Hucre_Fonksiyonlari.md)
- [Okunabilir Türkçe/İngilizce metinler](07_Documentation/V1_Okunabilir_Metinler.md)
- [Performans incelemesi ve kalan sınırlar](07_Documentation/Performans_Inceleme_Raporu.md)

## Geliştirici

**HAMU — Hüseyin Avni UZUN**

- [huseyinavniuzun.com](https://huseyinavniuzun.com)
- [GitHub](https://github.com/huseyinavniuzun)
- [hamu@huseyinavniuzun.com](mailto:hamu@huseyinavniuzun.com)

Bu depoda ayrı bir açık kaynak lisansı henüz belirtilmemiştir.
