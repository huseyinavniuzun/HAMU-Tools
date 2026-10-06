# HAMU Tools V1 — Performans incelemesi

Tarih: 05.10.2026. İncelenen dosya: `08_Addin/HAMUToolsV1.xlam`, kullanıcının yeni HAMU Tools klasöründeki güncel sürümü.

## Sonuç ve kapsam

Gerçek bir takılma nedeni yeniden üretildi: Normal İfade ile değiştir işlemi, korumalı hücreye yazarken 1004 hatası aldığında Excel'in `ScreenUpdating` ve `EnableEvents` ayarlarını kapalı bırakıyordu. Bu, işlem sona erse bile ekranın yenilenmemesine ve Excel olaylarının çalışmamasına neden olabilir. Hata yolu ve ayarların geri yüklenmesi düzeltildi.

Eklentinin tamamındaki kodlar; açılış işlemleri, olaylar, zamanlayıcılar, genel Excel ayarları, kalıcı nesneler, aralık işlemleri ve çalışma sayfası fonksiyonları bakımından tarandı. Seçili riskli yollar izole Excel örneğinde çalıştırıldı. Her özellik her veri kümesiyle test edilmedi. Kullanıcının takıldığı canlı Excel oturumuna ait CPU kaydı veya döküm alınmadı; bütün donmaların tek nedeninin bulunduğu iddia edilmiyor.

Güncel dosyanız esas alındı. Ribbon XML'i, ikonlar, çalışma kitabı parçaları ve form tasarımları korunmuştur. Paket karşılaştırmasında değişen tek içerik `xl/vbaProject.bin` oldu. Takvim ve Normal İfade formlarının boyutları ve kontrol konumları ayrıca karşılaştırıldı.

## Kanıtlanan sorunlar ve uygulanan düzeltmeler

| Alan | Bulgular | Düzenleme |
|---|---|---|
| Normal İfade hata çıkışı | Korumalı hücre hatasından sonra ekran yenileme ve olaylar kapalı kalıyordu. | Başarı, hata ve iptal yollarında önceki Excel ayarları geri yükleniyor. |
| Normal İfade metin çıktısı | Metinlerin Excel tarafından yorumlanmasına karşı çıktı türü açıkça korunmuyordu. | Telefon ve ayıklanan metinler `@` metin biçimi ve `Value2` ile yazılıyor. `+90` ve `=2+2` örnekleri formüle dönüşmeden test edildi. Eşleşme sayısı sayısal kalıyor. |
| Genel durum yönetimi | Yakalanmış ayarları olan komutlarda 62 satır, olayları/ekranı önceki değerden bağımsız açıyordu. | Bu satırlar yakalanan önceki değerleri kullanıyor. Ortak durum sınıfı yalnızca değişmiş ayarları geri yazıyor. |
| Tekrarlanan açılış işlemleri | Ayarlar ve fonksiyon tanımları açılış/Ribbon yenileme yollarında tekrar hazırlanabiliyordu. | Ayarlar bellekte bir kez okunuyor; açık ayar yükleme isteği korunuyor. Fonksiyon açıklamaları ihtiyaç halinde ve dil başına bir kez kaydediliyor. Açılışta gereksiz Ribbon invalidation kaldırıldı. |
| Normal İfade hücre dolaşımı | Hücre başına Excel nesnesine erişim ve çok büyük kaynak alanları işlem yükünü artırıyordu. | Kaynak dizi olarak okunuyor. Tüm sütun seçimleri gerçek veriyle sınırlandırılıyor. İşlem başına üst sınır 50.000 hücre; aşımda sessiz kırpma yerine açıklama veriliyor. |
| Normal İfade tekrar giriş | Aynı araç birden fazla modeless form ve iç içe işlem oluşturabiliyordu. | Açık form tekrar kullanılıyor; çalışırken ikinci komut engelleniyor. İşlem sırasında form kapatma/iptal isteği güvenli kontrol noktalarında işleniyor. |
| Takvim nesneleri | Her takvim çiziminde 42 olay nesnesi yeniden oluşturuluyor, form ve düğmeler karşılıklı referans tutuyordu. | Aynı 42 nesne kullanılıyor; kapanışta form ve düğme referansları ayrılıyor. |
| Aktarma verileri | Seçim listesi kapanış sonrasında modül dizisinde kalabiliyordu. | Kapanışta liste boşaltılıyor. |
| İngilizce çeviri belleği | İngilizce kullanıldıktan sonra Türkçeye dönüldüğünde sözlük kalıyordu. | Türkçeye dönüşte İngilizce sözlük serbest bırakılıyor. İngilizce gerektiğinde yeniden oluşturuluyor. |
| Vurgu geçmişi | Eski dolgu değerleri sınırsız büyüyebilir ve kapanmış dosyaların kayıtları kalabilirdi. | 50.000 kayıt sınırı var; Vurguyu Kaldır ile boşalan sözlük bırakılıyor. Kapanmış kitapların kayıtları bir sonraki HAMU komutunda temizleniyor. Açık kitap listesi değişmemişse geçmiş yeniden taranmıyor. |
| Renk fonksiyonları | Örnek hücrenin rengi döngünün her adımında yeniden okunuyordu. | Örnek rengi çağrı başına bir kez okunuyor. Mevcut 250.000 hücre sınırı ve volatile olmama davranışı korunuyor. |
| Tablodan veri getirme | Arama sütunu her satır için ayrı Excel erişimiyle okunuyordu. | Arama sütunu bir kez diziye alınıyor. İlk eşleşme ve büyük/küçük harf davranışı korunuyor; kalıcı veri önbelleği eklenmedi. |
| Çıktı alanı kontrolü | Birleştirilmiş hücre kontrolü hedefte hücre hücre yapılıyordu. | Alan düzeyinde kontrol kullanılıyor; koruma, çakışma ve dolu hedef uyarıları korunuyor. |
| Kullanılmayan kod | Eski dinamik Ribbon etiket modülündeki 155 callback artık XML veya diğer kodlar tarafından çağrılmıyordu. | `modHAMU_RibbonLabels` kaldırıldı. Kullanılan özellikler ve HAMU hücre fonksiyonları kaldırılmadı. |
| Durum çubuğu | Bu Excel ortamında False ataması bazı bağlamlarda metin `FALSE` olarak görünüyordu. | Ortak geri yükleme kodu bu durumu denetleyip metni temizliyor; önceden var olan özel durum mesajını koruyor. |

Telefon ayıklamanın kullanıcının özgün verisinde formül yazması henüz birebir yeniden üretilemedi: eski sürümde basit `+90` örneği zaten metin çıktı. Yeni sürümün açık metin yazması ayrı testlerle doğrulandı. Özgün hatalı sonuç ve kullanılan şablon, kalan olasılıkları ayırt etmek için gereklidir.

## Eklenti kullanılmıyorken neler çalışıyor?

- Sürekli `OnTime` zamanlayıcısı, `OnKey` kancası, uygulama düzeyinde hücre seçimi/değişimi/hesaplama izleyicisi veya `Application.Volatile` çağrısı bulunmadı.
- Eklentinin tek çalışma sayfasında hücre ve formül yok. Paket içinde bağlantı, externalLinks veya calcChain parçası bulunmadı.
- Açılışta yalnızca ayarlar yükleniyor. Ribbon görünürlük ve ikon isteklerine Excel gerektiğinde callback çağırabilir; bunlar veri taraması başlatmıyor.
- Takvim olayları yalnızca açık formun düğmelerine bağlıdır. Normal İfade kataloğu/form verileri araç kullanılınca hazırlanır.
- Kod, ikonlar ve Excel'in yüklediği VBA projesinin normal bir bellek maliyeti vardır. Kullanılmayan her özelliğin sıfır bayt tükettiği söylenemez; sürekli işlem yapmakla dosyada kod bulunması farklıdır.
- Çalışma kitabında HAMU hücre formülleri kullanılıyorsa, ilgili formüller Excel yeniden hesaplama yaptığında çalışabilir. Ribbon'a basılmaması bu formüllerin hesaplanmayacağı anlamına gelmez.

## Kullanım sırasında kalan yükler ve sınırlar

| İşlem | Kalan risk / sınır |
|---|---|
| Çok sayıda renk formülü | Renk biçimi hücre hücre okunmak zorundadır. VBA UDF hesaplamaları Excel'in çok iş parçacıklı yerleşik fonksiyonları gibi çalışmaz. Çok sayıda geniş alan formülü, dosya açılışını ve yeniden hesaplamayı yavaşlatabilir. |
| Özel Normal İfade | VBScript.RegExp içinde gerçek zaman aşımı yoktur. İptal kontrolü 256 hücrede bir ve sonuç yazmadan önce yapılır; tek bir ağır desen çağrısını veya başlamış sonuç yazımını anında kesemez. Patolojik bir desen uzun metinde hâlâ takılabilir. |
| Toplu görsel, barkod/karekod | Çok sayıda şekil eklemek; görsel dosyası/servis beklemek ve kitabın görsel yükü işlem sırasında yavaşlatabilir. Bu özellikler kendiliğinden çalıştırılmıyor. |
| Klasör tarama, dosya toplama, yedekleme | Dosya sayısı, OneDrive eşitlemesi, ağ/disk hızı ve dış dosyaların açılması süreyi etkiler. Bu işlemler komut verildiğinde çalışır. |
| Veri birleştirme, karşılaştırma, kalite ve vurgulama | Büyük alanlar için veri miktarına bağlı işlem maliyeti vardır. Bütün algoritmalar için aynı hız garantisi yoktur. Vurgu geçmişi sınırına ulaşılırsa önceki vurgular kaldırılabilir; işlem o ana kadar uyguladığı vurguları bırakabilir. |
| Başka eklentiler / dosyanın kendi kodu | Bu inceleme HAMU projesini kapsıyor. Diğer eklentiler, Power Query, dış bağlantılar ve açık kitapların olay kodları ölçülmedi. HAMU kullanılmadan süren genel takılma tamamen çözülmüş kabul edilmemelidir. |

## Kontroller

VBA son hali derlendi; geçici test modülleri ve test erişim yordamları teslim dosyasında bulunmuyor. Doğrulananlar: 5.000 telefon çıktısı, `+90` korunması, formül görünümlü metnin metin kalması, sayısal eşleşme sayısı, 50.000 hücre sınırı, farklı sayfaya değiştirme, kaynak formüllerini koruma, korumalı hedefte ayarları geri yükleme, çağıranın kapalı ekran/olay ayarlarını koruma, manuel hesaplama modunu koruma, iptalde boş hedefe sonuç yazmama, ikinci işlem engeli, tek Normal İfade formu, 42 takvim nesnesini yeniden kullanma ve referanslarını bırakma, aktarım/çeviri/vurgu belleği temizliği, tablo araması ve renk sayımı.

Son izole 5.000 hücre testinde eski sürüm 0,085938 sn, yeni sürüm 0,070313 sn ölçüldü. Bu küçük tek test genel Excel hızlanma oranı değildir; milisaniyelik ölçümler değişebilir.

## Kurulum ve geri dönüş

Güncellenen dosya aynı ad ve aynı `08_Addin` konumunda teslim edilir. Önceki eklenti ve değiştirilen kaynak dosyalar `09_Backups` altında tarih/saatli klasöre kopyalanır. Açık Excel oturumu eski VBA projesini bellekte tutabilir; bütün çalışmalar kaydedildikten sonra Excel tamamen kapatılıp yeniden açılmalıdır. Bu işlem kullanıcının Excel oturumunu otomatik kapatmaz.

Kaynaklar VBA ile uyumlu Windows-1254 kodlamasında ve BOM olmadan tutulur. Türkçe metinler hex'e çevrilmedi. Kullanıcının form ve Ribbon düzenleri yeniden tasarlanmadı.

## Teknik dayanak

[Microsoft — Excel performansını iyileştirme](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-tips-for-optimizing-performance-obstructions): diziyle toplu okuma/yazma, uygulama durumunu saklayıp geri yükleme ve VBA UDF maliyetleri.

[Microsoft — Excel yeniden hesaplama](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-recalculation): formüllerin yeniden hesaplama davranışı. [Microsoft — StatusBar](https://learn.microsoft.com/en-us/office/vba/api/excel.application.statusbar): durum çubuğu sahipliğinin Excel'e bırakılması.
