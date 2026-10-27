# HAMU Tools V1 — Açılış ve Markdown güncellemesi

14.10.2026

## Gerçekte yüklenen dosya

Excel'in kayıtlı açılış girdisi `"HAMUToolsV1.xlam"`. Bu adla yüklenen kopya `%APPDATA%\Microsoft\AddIns\HAMUToolsV1.xlam` konumundaydı. Masaüstündeki geliştirme kopyasıyla aynı değildi: aktif kopyada bazı genel veri/aktarım sınırları 1.000.000 hücreye yükseltilmişti. Güncelleme aktif kopya esas alınarak hazırlandı; bu değişiklikler korundu. Hem aktif dosya hem geliştirme dosyası ve kaynakları aynı sürüme getirilecek.

## Açılışta önceden ne yapıyordu?

- `Workbook_Open` içinde ayar yükleme çağrısı vardı.
- `RibbonOnLoad` da ayar yüklemeyi çağırıyordu. Önbellek mevcutsa ikinci dosya okuması yapılmıyordu; ancak ilk yükleme açılışa bağlıydı.
- Dil görünürlüğü ve ayar özeti callback'leri ayar sözlüğüne erişiyordu. Özet, boş yedek yolu için Windows Belgeler klasörünü de sorguluyordu.
- Ayar yükleme bir Dictionary ve FileSystemObject oluşturuyor, varsa UTF-16 `settings.ini` dosyasını okuyordu. Bu bir veri taraması veya sürekli dinleme değildi.
- İşlem sonunda ayar/çeviri sözlükleri her durumda bırakılmıyordu. Hakkında ekranı varsayılan form örneğini kullanıyordu.

## Uygulanan değişiklikler

1. `Workbook_Open` kaldırıldı. Açılışta iş komutu, dosya/klasör taraması, form oluşturma ve fonksiyon kayıt işlemi yok.
2. Ribbon açılışı yalnızca Ribbon nesne referansını tutuyor ve küçük dil tercihini bir kez okuyor. Görünürlük callback'leri Boolean/String durumunu kullanıyor; ayar dosyası, çeviri sözlüğü veya veri aralığı açmıyor. Excel'in Ribbon'u çizmek için bu callback'leri çağırması normaldir; bunlar hücre izleyicisi değildir.
3. Dosya menüsündeki dinamik yedek yolu özeti statik açıklamaya çevrildi. Gerçek konumlar Ayarlar ekranında görülür.
4. Ayar dosyası yalnızca araç ihtiyaç duyduğunda yükleniyor. Okuma 32 KB / 64 satırla sınırlandı; ayar dosyasının büyümesi açılışı etkileyemez.
5. Komut tamamlandığında/hata çıkışında geçici ayar ve İngilizce çeviri sözlükleri bırakılıyor. Dil değişimi Ribbon'a canlı yansıyor; kaydedilmiş dil açılış için küçük ayrı tercihte tutuluyor.
6. Hakkında formu yerel bir örnek olarak açılıp kapanınca Unload ediliyor. Markdown dosya akışı ve geçici tablolar işlem sonunda kapanıyor/bırakılıyor.
7. Araçlar yalnızca çağrıldığında çalışır. Takvimin form içi tıklama nesneleri, açık Normal İfade formu ve mevcut Vurguyu Kaldır geçmişi kendi kullanım ömrüne bağlıdır. Vurgu geçmişi ilk rengi geri getirmek için saklanan sınırlı veridir; arka plan dinleyicisi değildir.

Yeni zamanlayıcı, `OnKey`, uygulama düzeyinde hücre/seçim/değişim/hesaplama dinleyicisi veya volatile fonksiyon eklenmedi. Tek WithEvents nesnesi açık takvim formunun gün düğmesine aittir. Klasörü Aç için ShellExecute yalnızca o komutta çağrılır.

## Çökme kayıtları ve kanıtın sınırı

10–11 Ekim Windows Application kayıtlarında Excel `16.0.20430.20140` için `0xc0000005` erişim ihlalleri ve `OFFICE_MODULE_VERSION_MISMATCH` raporları görüldü. Hatalı modül `EXCEL.EXE` olarak kaydedilmiş. Bu, HAMU VBA'sında hangi satırın çökmeyi başlattığını göstermiyor ve Office sürüm uyumsuzluğu etiketinin tek başına kesin kök neden olduğunu kanıtlamıyor. Aralıklı yerel çökmeler izole testte yeniden üretilemedi. Bu nedenle “tüm çökmeler kesin çözüldü” iddiası yok.

Eklentinin sayfalarında hücre/formül bulunmadı. HAMU hücre formülleri kullanıcı çalışma kitaplarında kullanılıyorsa Excel yeniden hesaplamada bunları çalıştırabilir; Ribbon'a basılmaması bunların duracağı anlamına gelmez. Kullanıcı formüllerini bozacak şekilde bu özellik devre dışı bırakılmadı.

## Boşta ölçüm

Üç ayrı yeni Excel örneğinde güncel paket olaylar açıkken açıldı; çökme görülmedi. Eklenti açma süreleri yaklaşık 0,80–2,64 saniye aralığındaydı. Bu süre dosyayı açma/VBA ve arayüz yükleme dahil Excel maliyetidir; açılış iş yordamı çalışması değildir.

Ek karşılaştırmada her senaryoya boş kitap eklendi, 5 saniye başlangıç beklemesi sonrası 8 saniyelik süreç CPU süresi ölçüldü:

| Senaryo | 8 saniyede süreç CPU süresi | Özel bellek değişimi |
|---|---:|---:|
| Yalnız Excel | 0,328 sn | −4,04 MB |
| Önceki aktif HAMU | 0,078 sn | −3,88 MB |
| Yeni HAMU | 0,063 sn | −0,11 MB |

Bu birer kısa örnektir; hızlanma oranı veya bellek sızıntısı ölçümü değildir. İlk birkaç saniyede tüm Excel sürecindeki başlangıç işi daha yüksek CPU kullanıyordu. Ölçüm HAMU komutu çağrılmadan sürekli bir HAMU işi göstermez; kullanıcı dosyaları, başka eklentiler veya uzun oturumlar hakkında garanti vermez.

## Markdown aktarımı

**HAMU Tools > Veri > Birleştir > Markdown Aktar**.

- `.md` / `.markdown`, UTF-8 ve en fazla 2 MB.
- Başlık + `---` ayırıcı satırı olan pipe tablolarını tanır; hizalama için `:---`, `---:` ve `:---:` desteklenir.
- Bir tablo varsa doğrudan; birden fazla varsa Tümü veya seçilen tablo aktarılır. Her tablo yeni Excel sayfasında oluşturulur.
- Kod blokları atlanır. `\|` hücre içinde `|` olarak korunur. Boş alanlar ve Türkçe metin desteklenir.
- Telefonlar, baştaki sıfırlar ve `=2+2` gibi metinler formül olmadan metin olarak aktarılır.
- Oluşturulan tablo `TableStyleMedium1` kullanır. Sütunlar sığdırılır, çok uzun metinde genişlik 60 ile sınırlanır.
- Toplam 50.000 hücre, tablo başına 256 sütun ve hücre başına 32.767 karakter sınırı vardır. Yanlış sütun sayısı açıklamayla reddedilir; sessiz veri kırpılması yapılmaz.
- Hata/iptal yazım sırasında oluşursa yalnızca bu aktarımın oluşturduğu sayfalar kaldırılır. Önceki Excel ekran, olay, uyarı ve hesaplama ayarları geri yüklenir.
- Genel Markdown yazıları/paragrafları veya HTML tablolarını ayrıştırmaz. Hücrelerdeki Markdown biçimleme metni korunur; görsel metin render işlemi yapılmaz.

## Doğrulama

Son VBA projesi derlendi. Testler: Ribbon'un ayar/çeviri/form hazırlamaması, komut sonunda önbellek bırakılması, canlı dil durumu, UTF-8 okuma, birden fazla tablo, iki tür kod çiti, kaçışlı ayraç, boş hücreler, bozuk satır ve hücre sınırı, tek tablo seçimi, metin çıktısı, tablo stili, korumalı kitap ve yarım aktarımı geri alma. Excel'in çağıran ekran/olay/hesaplama ayarları korundu. Test yordamlarını yayımladığım eklentiden çıkardım.

Paket karşılaştırmasında yalnızca VBA projesi ve Ribbon XML'i değişti. Mevcut ikonlar ve form tasarım kaynakları aktif kopyadan korunmuştur. Türkçe VBA kaynakları Windows-1254, XML ve belgeler UTF-8'dir.

## Sonraki kullanım

Önceki aktif ve masaüstü dosyaları tarihli yedeklenir. Güncel eklenti her iki konuma uygulanır. Açık Excel oturumları otomatik kapatılmaz. Kaydedip Excel'i tamamen kapatarak yeniden açın; eski VBA projesi açık süreç belleğinde kalabilir.

Çökmeler sürerse Office kayıtlarını da kapsayan sonraki adım eklentisiz/güvenli modda aynı dosyanın denenmesi ve Office onarımıyla ayırmadır. Office kurulumu bu çalışmada değiştirilmedi. [Microsoft'un Excel donma/çökme rehberi](https://support.microsoft.com/en-us/excel/excel-not-responding-hangs-freezes-or-stops-working) eklentileri ve açılış dosyalarını ayrı ayrı kontrol etmeyi ve Office onarımını anlatır. [Ribbon nesnesi belgesi](https://learn.microsoft.com/en-us/office/vba/api/office.iribbonui) açılış ve Ribbon callback davranışını açıklar.
