# En son yerleşim

# HAMU Tools V1 — Ribbon yerleşimi ve RibbonX düzenleme

- Hızlı / Tablo düğmesi iki satırdır.
- Karekod / Barkod tek düğme ve tek formdur. Kod Türü seçiminde Karekod ve Barkod bulunur; model listesi türe göre güncellenir. İçerik, içerik türü ve kaynak/hedef adresi korunur. Toplu giriş aynı ekrandadır.
- Sayfa İndeksi ve Toplu Fotoğraf, Sayfa ve Görsel grubunun başında büyük düğmelerdir.
- Sayfaları Karşılaştır, Veri > Ayıkla menüsünde sondan ikinci sıradadır.
- Sorunlu / Hücreler düğmesi, Denetim / Araçları menüsünün önündedir.
- Bilgi grubunda fx simgeli büyük Ek Fonksiyon menüsü vardır. Hücre Fonksiyonları ve LAMBDA Şablonları bu menüdedir.
- Dosya > HAMU Tools'un iki sütununda da bir boş satırla başlangıç yapılır. Office sürümüne göre gerçek çizim yüksekliği değişebilir.


---

# Güncel Ribbon ve PDF kullanımı



## Güncel değişiklikler

- Veri İşlemleri **Birleştir, Ayıkla, Dönüştür, Kalite** olmak üzere dört büyük menüye indirildi. Grupla, Unpivot ve Pivot aynı Dönüştür menüsündedir. Menü içindeki araçlar büyük simgelerle gösterilir.
- Dosya grubu tek **Dosya** menüsüdür: Yedek Oluştur, Tarihli Sürüm, Kopya Kaydet, Klasör İçeriği ve PDF Kaydet.
- **Kopya Kaydet** seçilen konuma mevcut biçimde kopya oluşturur. Açık dosyanın adı, konumu ve kaydedilmemiş değişiklik durumu korunur. Makrolar biçim değiştirilerek kaybedilmez. Kaynak dosyanın kendisi, mevcut bir hedef ve yanlış uzantı reddedilir. Kaydedilmemiş makrolu dosyanızı önce uygun biçimde kaydedin.
- PDF'de **Alan Seç** ve **Etkin Sayfa** seçenekleri vardır. Alan Seç her zaman Excel'in seçicisini açar; mevcut alan kutuya başlangıç değeri olarak gelir. İptal edilirse dosya kaydedilmez. Etkin Sayfa mevcut sayfanın gerçek veri alanını kullanır. Otomatik, dikey ve yatay yön seçenekleri korunur.
- **Kırp ve Temizle**, **Boşlukları Sil**, **Harf Dönüştür** adları kullanılır. Diğer uzun başlıklar da kısaltıldı; detaylar araç ipuçlarında ve rehberde bulunur.
- 94 araç simgesi yeniden çizildi: gri yüzey, büyük işlev şekli, belirgin çizgi ve sınırlı vurgu rengi. Tarihli Sürüm takvim simgesidir; Yedek Oluştur kayıt simgesi; Kopya Kaydet üst üste iki belge. Bilgi simgesinin i gövdesi ve noktası yatay/dikey ortalanır.
- 16/32/64 px PNG, SVG ve 256 px PNG master kaynakları birlikte verilir. Master PNG'ler 1024 px'den küçültüldü; SVG ölçeklenebilir kalır. Logo hatları değiştirilmedi.


---

# Güncel V1 kullanım notu



## Güncel kullanım

- **Karekod / Barkod:** Kod İçeriği alanının altında Alan Seç ile tek sütunlu bir kaynak seçin. Dolu, hatasız hücreler için görseller aynı satırın sağındaki hücreye yerleştirilir; boş ve hata hücreleri atlanır. Sağdaki hücrelerde veri varsa işlem başlamaz. Aynı HAMU kod görseli yeniden üretildiğinde güncellenir. Metin olarak saklanmış 00123 gibi değerler korunur. Model ve içerik türü seçimleri toplu işlemde de kullanılır. Alanı Temizle ile tek kod girişine dönülür. İnternet gerekir ve kod içeriği seçilen dış servise gönderilir.
- Toplu kod üretiminde en fazla **250 kod** oluşturulur. Hücreler görseli barındıracak şekilde genişletilir; ilgili satır yüksekliği ve hedef sütun genişliği de değişir. İşlem kısmen başarısız olursa oluşturulanlar korunur, hata sayısı ve ilk hata adresi gösterilir. Esc ile kesilen işlem Excel durumunu geri yükler.
- **Veri Girişi:** Tablo veya başlıklı veri listesinin içindeyken Excel'in kendi veri giriş ekranını açar. En fazla 32 sütun ve boş olmayan, benzersiz başlıklar gerekir. Yeni bir HAMU formu oluşturulmadı.
- **Kolon Seç:** Başlık adları veya seçim içindeki kolon numaraları kullanılabilir: `1,2,3`. A:G gibi seçimler son dolu satıra kadar daraltılır. Yalnız biçim verilmiş uzak hücreler veri kabul edilmez; boş sonuç veren formüller korunur. Geçersiz veya çok büyük kolon numaraları kopyalama başlamadan reddedilir. Çıktı gerçek bir Excel tablosudur.
- Hücre hücre çalışan ortak seçim yolları ve doğrudan kullanılan büyük veri taramaları **500.000 hücre** sınırıyla korunur. Bu sınır bir performans garantisi değildir; karmaşık işlemlerde daha küçük alanlar seçin.
- **Resim Kaydet:** Seçili alan PNG, JPG veya GIF olarak kaydedilir. Geçici grafik temizlenir. Excel'in resmi çizmesi beklenir ve kopyalama geçici olarak başarısız olursa yeniden denenir. En fazla 100.000 hücre ve her yönde 10.000 punto boyut kabul edilir.
- **PDF Kaydet:** Seçili aralık veya aktif sayfa; otomatik yön, dikey veya yatay seçilebilir. Otomatik yön alanın en-boy oranına bakar. Geçici baskı alanı, yön ve ölçek ayarları işlem sonunda geri yüklenir.
- **Vurguyu Temizle:** Seçili alandaki HAMU denetim vurgularını kaldırır. Aynı oturumda önceki dolgu geri yüklenir. Excel yeniden açıldıktan sonra önceki dolgu bilinmez; HAMU'nun sarı vurgu renkleri kaldırılır. Aynı renkteki elle yapılmış dolgular da bu ikinci durumda etkilenebilir.
- **Kullanım Rehberi:** Ara kutusu konu metinlerinde ve 93 araç açıklamasında arar. Aramayı temizleyince seçili konuya dönülür. Türkçe harfler ve büyük/küçük harf farkı aramayı engellemez.
- **Başlıkları Düzelt:** Büyük İ / küçük ı / I harfleri snake_case dönüşümünde i olur; İŞLEM → islem. Temiz başlık ve snake_case seçenekleri seçim listesinde gösterilir.

## Menü düzeni

Dosya Yönetimi: büyük **Yedek Oluştur** ve **Tarihli Sürüm**; doğrudan **Klasör İçeriği** ve **PDF Kaydet**. Yenile kaldırıldı.

Tablo ve Veri Girişi: **Tablo Oluştur**, **Veri Girişi**, **Kolon Seç** ve tablo araçları. **Tabloyu Genişlet** tablo araçları menüsündedir. **Sığdır** sütun genişliği ve satır yüksekliğini düzenler; tabloya yeni satır/sütun katmaz. Bu nedenle Metin ve Hücre grubundaki hücre düzeni araçlarına taşındı.

Kontrol ve Denetim menüsünde **Formül ve Ad Denetimi** ilk sıradadır. Vurguyu Temizle aynı gruptadır. Yardım ve Hakkında tek büyük **Bilgi** menüsünde toplanır; hücre fonksiyonları ve LAMBDA açıklamaları buradan açılır. LAMBDA araçlarında ortak λ ikonu kullanılır.

## Aktarma Sihirbazı

Gönderilen orijinal 64 px HAMU logosu kullanılır. Aktarım türleri: kitap kapsamlı LAMBDA, özel hücre stili, görünür adlandırılmış veri aralığı, sayfa verisi ve Excel tablosu.

Sayfa/tablo/aralık aktarımı değerleri, sayı biçimlerini, hücre biçimlerini ve sütun genişliklerini kopyalar. VBA ve hücre formülleri taşınmaz; kaynak dosyaya bağlanan yeni formüller oluşturulmaz. Sayfa ve tabloda çakışma için atla veya yeniden adlandır kullanılır. Mevcut hedef sayfa silinmez. Kullanıcının önceden açtığı kaynak kitap kapatılmaz.

## Gerçek Excel LAMBDA şablonu

**Bilgi > LAMBDA Şablonu > Şablonu Aç** yolundan `09_Templates/HAMU_LAMBDA_Sablonu.xlsx` açılır. Aktarma Sihirbazı ile tanımları çalışma kitabınıza aktarabilirsiniz. Bunlar VBA yardımcı fonksiyonu değildir; Excel'in Ad Yöneticisi'nde görünür.

1. `HAMU_L_CokluAra`: iki anahtar/sonuç aralığı çiftinde sırayla tam eşleşme arar; ilk sonucu verir. Bulamazsa boş döner. Tek sütunlu ve eşit uzunlukta eşleşme çiftleri gerekir.
2. `HAMU_L_CokluTopla`: iki anahtar/değer aralığındaki aynı anahtarın sayısal karşılıklarını toplar.
3. `HAMU_L_BenzersizBirlestir`: iki aralığı tek listeye birleştirir, boşları kaldırır, benzersiz değerleri sıralar.

Örnekler şablonda hesaplanmış halde bulunur. Bu şablon LAMBDA/VSTACK/TOCOL gibi güncel Excel işlevlerini gerektirir; güncel Microsoft 365 veya Excel 2024 kullanın. Eski Excel sürümlerinde çalışmayabilir. Excel'in kendi gizli `_xlfn` / `_xlpm` uyumluluk adları silinmez. Eklentinin iç yardımcı modülleri Option Private Module ile gizlidir; sekiz HAMU VBA hücre fonksiyonu kullanılmaya devam eder.



---

# HAMU Tools V1 — Kullanım rehberi

## İlk adımlar

HAMU Tools, Excel içinde dosya, veri ve hücre işlemlerini hızlandırır.

1. İşlem yapacağınız kitabı ve sayfayı açın.
2. Kaynak veriyi seçin; başlık isteyen araçlarda başlık satırını da ekleyin.
3. Şeritten uygun aracı açın, açıklamasını okuyun ve seçenekleri belirleyin.
4. Yeni çıktı isteyen araçlarda boş bir hedef seçin; sonuç ve satır sayısını kontrol edin.

Önerilen akış: Yedek Al → Başlıkları Normalize Et → Veri Kalitesi → Tablo Oluştur → Grupla veya Unpivot.

Komutların tamamı şerittedir. Dosya > HAMU Tools ise iş akışlarını ve bu rehberi sunar. Çalışma kitabı açık değilken yardım ve hakkında ekranlarına erişebilirsiniz.

## Aralık ve çıktı seçimi

Mevcut seçiminiz, aralık kutusunda başlangıç adresi olarak görünür. Kutudan farklı hücreleri seçebilir veya geçerli bir Excel adresi yazabilirsiniz. Seçimi Enter ile onaylayın; X ile kapatmak işlemi iptal eder.

Kaynak, okunacak veridir; hedef ise sonucun yazılacağı yerdir. Birden fazla giriş isteyen araçlarda kutu başlığını kontrol edin. Tek hedef hücresi, liste sonuçlarında başlangıç noktasıdır; sonuç aşağıya veya sağa genişleyebilir.

Başlıklar: Tablo, birleştirme, grupla ve pivot araçlarında ilk satırın anlamlı ve benzersiz başlıklar taşıması gerekir.

Tarih Seç, önceden seçili aktif hücreye yazar. Formdaki ay ve yılı değiştirebilir, Bugün ile güncel tarihe dönebilirsiniz. Seçilen tarihin bugüne uzaklığı altta gösterilir.

## Metin ve veri temizleme

Önce küçük bir örnekte deneyin, ardından tüm listeye uygulayın.

Başlıkları Normalize Et: Boşluk ve yinelenen başlıkları düzenler. Metin dönüşümleri: Büyük/küçük harf, boşluk, ön/son ek ve metin parçası araçlarını kullanın. Normal İfade: Şablon seçin veya özel desen yazın; açıklama ve örnekleri inceleyin.

Boşlukları Sil, boş hücreleri kaldırıp alttaki hücreleri yukarı taşır; satırlar arasında ilişkili sütunlar varsa sonuçları dikkatle kontrol edin. Boş hücreleri sıfırla doldurmak hücre silmez.

Standart Sayı, tam sayı görüntüsü uygular; Standart Tarih, tarih görünümünü düzenler. Hücrenin gerçek türü ile görüntü biçimi farklıdır. Baştaki Sıfırlar, kodları metin olarak saklar; hesaplama yapılacak sayılarda kullanmayın.

Veri Kalitesi araçlarıyla boşlukları, hataları ve tekrarları inceleyin; denetim sonucu bir temizleme kararına yardımcı olur.

## Tablolar ve özetleme

Tablo Oluştur: Başlıkları içeren veri alanından Excel tablosu oluşturur. Otomatik Sığdır: Satır ve sütun boyutlarını içeriğe göre düzenler. Tabloyu Genişlet: İmleç bir Excel tablosunun içindeyken kullanılır.

Kolon Seçerek Yeni Tablo: İhtiyacınız olan sütunlardan ayrı çıktı üretir. Grupla: Bir veya daha fazla anahtar sütuna göre özet oluşturur. Pivot: Kategorileri sütunlara taşır. Unpivot: Geniş sütun yapısını alan/değer satırlarına dönüştürür; 2026.01 gibi metin başlıkları korunur.

Akıllı Birleştir aynı yapılı listeleri alt alta toplar. Klasörlerden Veri Topla birden fazla dosyayı okur; dosya ve sayfa yapısını önceden kontrol edin. Anahtarla Eşleştir, ortak anahtara göre ayrı listeleri ilişkilendirir. Anahtarların türü, boşlukları ve tekrarları eşleşmeyi etkiler.

## Dosyalar ve yedekleme

Yedek Al, açık dosyanın tarih ve saatli kopyasını Belgeler > HAMU > Backup klasörüne kaydeder. Sonuç ekranından tam yolu kopyalayabilir ve Klasörü Aç ile klasöre ulaşabilirsiniz.

Dosya Adına Tarih Ekle, kaydedilmiş dosyanın aynı konumda yeni kopyasını oluşturur; kopya açılıp doğrulandıktan sonra eski dosyayı kapatıp siler. Bu işlemi yedekten sonra kullanın.

Klasör İçeriğini Listele: Yalnız mevcut klasörü veya alt klasörleri de kapsayan taramayı seçin. Liste; bağlantı, yerel yol, tür, boyut ve oluşturma/değiştirme tarihlerini içerir. Çok sayıda alt klasör daha uzun sürebilir.

HAMU Transfer: Şablon dosyasındaki kitap kapsamlı LAMBDA adlarını ve hücre stillerini aktarır. Aynı adlı öğelerde atla, üzerine yaz veya yeniden adlandır seçeneklerini inceleyin. VBA işlemleri Excel geri alma geçmişini temizleyebilir; kritik verileri önceden yedekleyin.

## Resimler, barkod ve karekod

Barkod/Karekod ekranında hedef hücreyi ve içerik modelini seçin. Metin, bağlantı veya e-posta gibi seçenekler aynı içeriğin doğru kod biçiminde hazırlanmasını sağlar. E-posta seçimi ileti göndermez; kodu tarayan kişinin uygulaması bir taslak açabilir.

Kod üretimi internet ve üçüncü taraf servis kullanır; kod içeriği üretim servisine gönderilir. Gizli veya kişisel veri girmeden önce bunu değerlendirin. Bağlantı veya servis sorunu varsa kod üretilemeyebilir.

Klasörden Resim Ekle: Hücre içine veya hücre üstüne yerleştirme seçeneğini belirleyin. Hücre içine özellik desteği Excel sürümüne bağlıdır; hücre üstündeki resimler hücre sınırlarına göre yerleştirilir.

Satırları Renklendir ve sayfa/görsel araçları ilgili gruptadır. Çalışma sayfası renk formülleri doğrudan dolgu rengini okur; denetim komutları görüntülenen rengi, koşullu biçimlendirme dahil, kullanır.

## HAMU hücre fonksiyonları

Formülleri hücreye =HAMU_ yazarak bulun. Aşağıdaki örneklerde Türkçe Excel ayırıcıları kullanılmıştır; Excel diliniz farklıysa ayırıcıyı uyarlayın.

=HAMU_URLCoz(A1) | UTF-8 URL karakterlerini çözer. Geçersiz yüzde dizisini korur.

=HAMU_RengeGoreTopla(A1:A20;B1) | B1 ile aynı dolgu rengine sahip sayıları toplar. Biçim değişimi sonrası Ctrl+Alt+F9 kullanın.

=HAMU_RengeGoreSay(A1:A20;B1) | B1 ile aynı dolgu rengine sahip hücreleri sayar. Biçim değişimi sonrası Ctrl+Alt+F9 kullanın.

=HAMU_YerelDosyaYolu(A1) | Kişisel OneDrive URL yolunu yerel OneDrive ortam değişkeniyle eşler; diğer URL türlerini korur.

=HAMU_GizliLink(A1) | Hücrenin ilk köprü adresini getirir.

=HAMU_LinkParamDegeri(A1;"id") | URL içindeki sorgu parametresinin değerini getirir.

=HAMU_SayidanMetine(123) | Türkçe: yüz yirmi üç. 15 basamaktan küçük mutlak değerler; ondalık kısmı tek tek okur.

=HAMU_TablodanVeriGetir(A1;"Tablo1";"Kod";"Ad") | Formülün bulunduğu kitapta tablodan ilk eşleşen değeri getirir.

Renk değiştirdikten sonra Ctrl+Alt+F9 ile yeniden hesaplayın. Renk formülleri otomatik bir renk değişimi olayı izlemez. HAMU eklentisi kapalıysa bu özel fonksiyonlar hesaplanamaz; dosyayı paylaşacağınız kişinin de eklentiye ihtiyacı olabilir.

## Sorun giderme ve sınırlar

Komut pasif veya çalışmıyor: Açık bir çalışma kitabı ve uygun sayfa olduğundan emin olun. Hücre düzenleme modundan Enter/Esc ile çıkın.

Aralık geçersiz: Doğru kitap/sayfayı seçin; aracın tek alan, tek sütun veya başlık şartını kontrol edin. Korunan sayfalarda yazma işlemleri engellenebilir.

Tablo bulunamadı: İmleci Excel tablosunun içine taşıyın veya önce Tablo Oluştur çalıştırın.

Formül sonucu beklenmiyor: Sayı/metin türünü, başlık adlarını, anahtar tekrarlarını ve F9 hesaplamasını kontrol edin.

Kod/resim oluşturulamadı: İnternet, servis erişimi ve Excel sürüm desteğini kontrol edin.

Uzun klasör ve veri işlemlerinde Excel bekleyebilir; alt klasör kapsamını daraltın ve küçük örnekle başlayın. Sorun bildirirken araç adı, Excel sürümü, hata metni ve kişisel veri içermeyen bir örnek paylaşın.

HAMU Tools V1, masaüstü Excel VBA eklentisidir; web Excel VBA çalıştırmaz.

## Tüm araçlar

Ana İşlem / HAMU Transfer
Excel şablonundaki kitap kapsamlı LAMBDA tanımlarını ve özel hücre stillerini seçerek aktarır. Çakışmalarda atlama, üzerine yazma veya yeniden adlandırma sunar.

Dosya Yönetimi / Yedek Al
Dosyanın tarih ve saatli kopyasını Belgeler > HAMU > Backup klasörüne kaydeder. Sonuç ekranında tam yolu gösterir; Klasörü Aç ile yedek klasörüne ulaşabilirsiniz.

Dosya Yönetimi / Dosya Adına Tarih Ekle
Kaydedilmiş dosyanın aynı klasörde tarih ve saat eklenmiş kopyasını oluşturur. Yeni kopya açılıp doğrulandıktan sonra eski dosya kapatılır ve silinir. Yeni dosya açılamazsa eski dosya korunur.

Dosya Yönetimi / Sayfaları PDF Olarak Kaydet
Çalışma kitabındaki görünür sayfaları ayrı PDF dosyaları olarak seçtiğiniz klasöre kaydeder.

Dosya Yönetimi / Klasör Listesi
Klasör seçildikten sonra yalnız bu klasör veya alt klasörlerle birlikte tarama seçilir. Alt klasör taraması uzun sürebilir. Dosyaları bağlantı, yol, tür, okunabilir boyut ve tarihlerle listeler.

Dosya Yönetimi / Verileri Yenile
Çalışma kitabındaki bağlantıları, PivotTable önbelleklerini ve RefreshAll kapsamındaki Power Query/veri bağlantılarını yeniler.

Veri İşlemleri / CSV Dosyalarını Birleştir
Seçtiğiniz klasördeki CSV dosyalarını tek sayfada alt alta toplar; ilk dosyanın başlığını kullanır ve sonraki dosyalarda başlık satırını atlar.

Veri İşlemleri / Sayfaları Alt Alta Birleştir
Aktif çalışma kitabındaki sayfaların kullanılan alanlarını tek sayfada birleştirir ve kaynak sayfa bilgisini sonuç verisine ekler.

Veri İşlemleri / Seçili Tabloları Alt Alta Birleştir
Belirttiğiniz Excel tablolarının aynı yapıdaki satırlarını tek bir sonuç tablosunda alt alta toplar.

Veri İşlemleri / Başlıklara Göre Akıllı Birleştir
Aynı bilgileri içeren sayfaları alt alta birleştirir. Örneğin Ocak ve Şubat satışlarında sütun sırası farklı olsa da Ürün başlıkları eşleştirilir. Eksik sütunlar boş kalır, KaynakSayfa eklenir. Kaynak sayfalar değiştirilmez; ilk satırlar başlık olmalıdır.

Veri İşlemleri / Klasörden Veri Topla
Klasördeki Excel ve CSV dosyalarının ilk veri sayfalarını tek sonuçta alt alta toplar. İlk satır başlıktır; aynı başlıklar eşleştirilir. KaynakDosya eklenir, kaynak dosyalar değiştirilmez. Aynı tür tabloları içeren dosyalar kullanın.

Veri İşlemleri / Anahtara Göre Tablo Eşleştir
İki tablodaki kayıtları müşteri kodu veya ürün kodu gibi ortak alanlarla yan yana eşleştirir. LEFT: ilk tablonun tüm satırları; INNER: yalnız eşleşenler; FULL: iki tablonun tüm kayıtları. İkinci tablodan alınacak sütunlar ayrıca seçilir. Kaynaklar değiştirilmez.

Veri İşlemleri / Grupla ve Özetle
Bir veya birden fazla kolona göre kayıtları gruplar; birden fazla sayısal kolonu aynı anda toplar. Kolonları numarayla veya başlık adıyla seçebilirsiniz.

Veri İşlemleri / Kolonları Satırlara Dönüştür
Solda belirlediğiniz birden fazla sabit kolonu korur; sağa doğru uzanan pivot kolonlarını Alan ve Değer yapısında satırlara dönüştürür.

Veri İşlemleri / Veriyi Böl
Seçili tabloyu satır sayısına, eşit parçalara, tek veya çoklu kolon değerlerine göre ya da yıl, çeyrek, ay ve hafta bazında tarih grupları oluşturarak böler. Çıktıları yeni çalışma sayfalarına veya ayrı Excel dosyalarına aktarabilirsiniz.

Veri İşlemleri / Değere Göre Ayrı Dosyalar Oluştur
Seçtiğiniz kolonun benzersiz değerlerine göre veriyi filtreler ve her değer için ayrı bir Excel dosyası oluşturur.

Veri İşlemleri / URL Parametrelerini Kolonlara Ayır
URL'lerde soru işaretinden sonra bulunan sorgu parametrelerini algılar ve her parametreyi ayrı bir kolon halinde yeni sayfaya çıkarır.

Veri İşlemleri / Pivot Filtrelerini Ayrı Sayfalara Aktar
Seçili PivotTable'ın rapor filtresindeki değerleri sırayla uygular ve her filtre sonucu için ayrı çalışma sayfası oluşturur.

Veri İşlemleri / Satırlardan Çapraz Tablo Oluştur
Uzun formattaki veriyi tekrar geniş çapraz tabloya dönüştürür. Birden fazla sabit anahtar kolonu kullanabilir; yeni kolon başlıklarını seçilen alandan üretir ve aynı hücreye düşen kayıtları toplama, ilk değer veya sayım yöntemiyle birleştirir.

Metin ve Hücre / Boşlukları Aşağı Doldur
Seçili aralıktaki boş hücreleri aynı kolondaki en yakın üst dolu değerle doldurur. Pivot veya rapor çıktılarındaki grup etiketlerini satırlara yaymak için kullanışlıdır.

Veri İşlemleri / Tekrarları Birleştir
Tek veya çoklu anahtar kolonlarla tekrar eden kayıtları tek satırda birleştirir. Seçtiğiniz sayısal kolonları toplar, diğer alanlarda ilk dolu değeri korur ve kaç kaydın birleştiğini ayrıca yazar.

Veri İşlemleri / Başlıkları Standardize Et
Kolon başlıklarını teknik snake_case biçimine dönüştürebilir veya gereksiz boşlukları temizleyerek okunabilir başlık yapısını koruyabilir. Aynı başlık oluşursa otomatik olarak benzersiz ad üretir.

Veri İşlemleri / Kolon Veri Tipini Düzelt
Bir veya birden fazla kolonu metin, sayı, tarih, tam sayı veya yüzde tipine dönüştürür; dönüştürülemeyen hücreleri sayarak işlem sonucunu raporlar.

Tablo ve Görünüm / Kolon Seçerek Yeni Tablo Oluştur
Geniş bir veri kümesinden yalnızca ihtiyacınız olan kolonları seçtiğiniz sırayla yeni bir sayfaya çıkarır. Kolonları başlık adı veya numarasıyla belirleyebilirsiniz.

Veri İşlemleri / Koşulla Satır Filtrele
Koşula uyan satırları yeni sayfaya kopyalar; kaynak değişmez. Şehir=Bursa yalnız Bursa kayıtlarını, Satış>0 pozitif satışları, Ürün~TV içinde TV geçenleri seçer. Noktalı virgülle birleştirilen koşulların hepsi sağlanmalıdır.

Veri İşlemleri / Veriden Örnek Al
Tüm veri kümesinden rastgele N satır seçebilir veya belirlediğiniz grup kolonlarına göre her gruptan en fazla N rastgele kayıt alarak yeni bir örnek veri sayfası oluşturabilir.

Veri İşlemleri / Satırları Karıştır
Kaynak veriyi değiştirmeden, başlık satırını koruyarak veri satırlarını rastgele sırada yeni bir çalışma sayfasına aktarır.

Veri İşlemleri / İki Listeyi Karşılaştır
İki tek kolonlu listeyi karşılaştırır; yalnızca birinci listede, yalnızca ikinci listede veya her iki listede bulunan değerleri raporlar.

Veri İşlemleri / Benzersiz Liste
Tekrar eden değerleri bir kez listeler. Çıktı için mevcut sayfada veya başka sayfada başlangıç hücresi seçebilir ya da yeni sayfa oluşturabilirsiniz. Dolu hedefte onay ister; kaynak hata ve boş hücrelerini atlar.

Veri İşlemleri / Veri Profilini Çıkar
Seçtiğiniz başlıklı veri alanındaki sütunların doluluk, boşluk, benzersiz değer ve veri türlerini özetler. Sorunları görmek için bir rapor oluşturur; kaynak veriyi değiştirmez.

Veri İşlemleri / Veri Kalitesi Raporu
Başlıklı tabloyu denetleyerek boş alan, tekrar eden kayıt ve veri türü sorunlarını raporlar. Sonuç yeni sayfaya yazılır; sorunları kaynaktan otomatik silmez. Önce raporu inceleyin.

Veri İşlemleri / Eksik Kombinasyonları Bul
Mağaza, Ürün, Hafta gibi seçtiğiniz kolonların benzersiz değerlerinden tüm olası kombinasyonları oluşturur ve kaynak veri setinde bulunmayan kombinasyonları yeni bir sayfada listeler.

Metin ve Hücre / Metin ve Boşluk Temizle
Seçili metinlerde görünmeyen NBSP karakterlerini, sekmeleri, satır sonlarını, sıfır genişlikli boşlukları ve art arda gelen gereksiz boşlukları temizler. Metnin başı/sonundaki boşlukları kaldırır; formülleri değiştirmez.

Metin ve Hücre / Büyük / Küçük Harf Dönüştür
Seçili metinleri tamamı büyük, tamamı küçük veya baş harfleri büyük biçime dönüştürür.

Metin ve Hücre / Normal İfade ile Bul
Seçili hücrelerde RegEx normal ifade desenine uyan metinleri bulur ve eşleşen hücreleri vurgular. Sayı, kod, telefon veya belirli metin kalıplarını tespit etmek için kullanılabilir.

Metin ve Hücre / Normal İfade ile Değiştir
Seçili metin hücrelerinde RegEx normal ifadeleri kullanarak toplu bul/değiştir işlemi yapar. Gruplama ve $1 gibi geri başvurular kullanılabilir.

Metin ve Hücre / Metinden Sayıları Ayıkla
Metin içindeki rakam karakterlerini ayıklar ve sonucu hedef alana aktarır.

Metin ve Hücre / Metin Olarak Saklanan Sayıları Düzelt
Başında gizli tek tırnak bulunan veya Excel tarafından metin olarak saklanan sayısal değerleri gerçek sayıya dönüştürmeye çalışır.

Metin ve Hücre / Veriyi Maskele
Seçili metinlerin belirlediğiniz kadar başlangıç ve bitiş karakterini koruyup ortadaki karakterleri yıldızla gizler.

Metin ve Hücre / Formülleri Değere Dönüştür
Seçili hücrelerdeki formülleri mevcut sonuçlarıyla değiştirir. İşlem sonrasında hücrelerde formül yerine sabit değer kalır.

Metin ve Hücre / Birleşimleri Çöz ve Doldur
Birleştirilmiş hücreleri çözer ve birleşik alanın sol üst değerini çözülen hücrelerin tamamına yazar.

Metin ve Hücre / Standart Sayı ve Tarih Biçimi
Gerçek sayıları ondalıksız tam sayı görünümüyle, tarih hücrelerini gg.aa.yyyy biçiminde gösterir. Hücre değerlerini yuvarlamaz veya değiştirmez.

Metin ve Hücre / Boş Hücreleri Sil ve Yukarı Kaydır
Seçili alandaki boş hücreleri silerek alttaki hücreleri yukarı kaydırır. Veri hizasını etkileyebileceği için tek kolonlu listelerde kullanılması önerilir.

Sayfa ve Görsel / Eski Dropdown Nesnelerini Sil
Aktif sayfadaki eski Form Denetimi ve ActiveX açılır kutu nesnelerini siler. Hücrelerdeki veri doğrulama listeleri korunur.

Sayfa ve Görsel / Satırları Renklendir
Seçili aralıkta dönüşümlü satır renklendirmesi uygular. Çalıştırıldığında mavi, yeşil, sarı, turuncu, mor, gri ve pembe tonlarından birini seçebilirsiniz.

Tarih Araçları / Tarih Seçici
HAMU takvim penceresini açar. Aktif hücrede tarih varsa takvim o tarihle açılır; seçtiğiniz tarih hücreye gerçek Excel tarihi olarak dd.mm.yyyy biçiminde yazılır.

Tarih Araçları / Tarihleri Düzelt
03/10/2026, 2026-10-03, 20261003, 3 Ekim 2026 ve benzeri farklı tarih gösterimlerini gerçek Excel tarihine dönüştürür ve dd.mm.yyyy biçiminde standardize eder.

Tarih Araçları / Tarih Kontrolü
Seçili hücrelerdeki değerlerin geçerli bir tarih olup olmadığını kontrol eder. Geçerli tarihleri ve tanımlanamayan değerleri farklı renklerle işaretler.

Tarih Araçları / Tarih Oluştur
Ayrı kolonlarda bulunan yıl, ay ve gün değerlerini birleştirerek gerçek Excel tarihleri oluşturur. Ay alanında sayı veya Türkçe ay adı kullanılabilir.

Sayfa ve Görsel / Sayfa İndeksi
Görünür çalışma sayfalarını listeleyen ve her sayfaya tıklanabilir bağlantı veren bir İçindekiler sayfası oluşturur.

Sayfa ve Görsel / Sayfaları Alfabetik Sırala
Çalışma kitabındaki çalışma sayfalarını adlarına göre alfabetik olarak yeniden sıralar.

Sayfa ve Görsel / Tüm Gizlileri Göster
Aktif çalışma kitabındaki gizli çalışma sayfalarını, satırları ve sütunları görünür hale getirir.

Sayfa ve Görsel / Dinamik Veri Doğrulama Listesi
Seçili hücrelere liste tipi veri doğrulama uygular. Kaynak olarak adlandırılmış aralık, Tablo[Kolon], doğrudan aralık veya A1# biçimindeki dinamik taşan dizi kullanılabilir. Bir adlandırma A1# kaynağına bağlıysa taşan aralık büyüdükçe doğrulama listesi de dinamik kalır.

Sayfa ve Görsel / Seçili Alanı PNG Olarak Kaydet
Seçtiğiniz hücre aralığını görüntü olarak kopyalar ve PNG dosyası şeklinde kaydeder.

Sayfa ve Görsel / Klasörden Toplu Fotoğraf Ekle
Seçtiğiniz klasördeki görselleri aktif hücreden başlayarak satırlara ekler ve yan hücreye dosya adını yazar.

Sayfa ve Görsel / Barkod Oluştur
Seçilen hücreye Code 128 barkod görseli ekler. Girilen değer TEC-IT internet servisine gönderilir; internet bağlantısı gerekir.

Sayfa ve Görsel / QR Kod Oluştur
Seçilen hücreye QR görseli ekler. Girilen metin/URL QuickChart internet servisine gönderilir; internet bağlantısı gerekir.

Kontrol ve Denetim / Sayfaları Karşılaştır
İki çalışma sayfasını belirlediğiniz anahtar kolona göre karşılaştırır ve eklenen, silinen veya değişen kayıtları ayrı bir raporda gösterir.

Kontrol ve Denetim / Sorunlu Hücreleri İşaretle
Sayı gibi görünen metinleri ve eşittir işaretiyle başlayan ancak gerçek formül olmayan metinleri bularak hücreleri farklı renklerle işaretler.

Kontrol ve Denetim / Hücre Türüne Göre Vurgula
Seçili alanda sayı, metin, hata veya boş hücreleri seçtiğiniz kritere göre vurgular.

Kontrol ve Denetim / Tekrar Edenleri Vurgula
Seçili alanda birden fazla kez bulunan değerleri tespit eder ve tekrar eden hücreleri renklendirir.

Kontrol ve Denetim / Tekrar Sayısını Yaz
Giriş listesindeki her değerin kaç kez geçtiğini hesaplayıp ayrı seçtiğiniz çıkış alanına yazar. Tek çıkış hücresi seçerseniz sonuç aşağı doğru genişler.

Kontrol ve Denetim / Harici Bağlantıları Listele
Çalışma kitabındaki Excel türü dış bağlantıları yeni bir sayfada listeler.

Kontrol ve Denetim / Harici Bağlantıları Kır
Çalışma kitabındaki harici Excel bağlantılarını kırar. Bağlantılı formüller mevcut değerlerine dönüştürülebileceği için işlemden önce yedek alınması önerilir.

Yardım ve Hakkında / Yardım
HAMU Tools içindeki komutların kullanım amacı ve temel kullanım bilgilerini gösterir.

Yardım ve Hakkında / Hakkında
HAMU Tools V1 sürümü, HAMU markası ve huseyinavniuzun.com bilgileri ile eklentinin kapsamını gösterir.

Tablo ve Görünüm / Otomatik Sığdır
Seçilen sütunları en fazla 60 karakter genişliğine sığdırır; uzun metinleri kaydırır ve satır yüksekliğini ayarlar.

Tablo ve Görünüm / Otomatik Tablo Oluştur
Tek hücre seçildiyse çevresindeki kesintisiz veri alanını, aksi halde seçilen alanı başlıklı Excel tablosuna dönüştürür. HAMU stili ve filtreleri uygular.

Tablo ve Görünüm / Tabloyu Veriye Genişlet
Seçili tablonun altındaki kesintisiz veriyi aynı sütun genişliğinde tabloya dahil eder. Başka tablolarla çakışmayı önler.

Yardım ve Hakkında / Çalışma Sayfası Fonksiyonları
HAMU çalışma sayfası fonksiyonlarını ve formül örneklerini gösterir.

Yardım ve Hakkında / HAMU_URLCoz
=HAMU_URLCoz(A1) | UTF-8 URL karakterlerini çözer. Geçersiz yüzde dizisini korur.

Yardım ve Hakkında / HAMU_RengeGoreTopla
=HAMU_RengeGoreTopla(A1:A20;B1) | B1 ile aynı dolgu rengine sahip sayıları toplar. Biçim değişimi sonrası Ctrl+Alt+F9 kullanın.

Yardım ve Hakkında / HAMU_RengeGoreSay
=HAMU_RengeGoreSay(A1:A20;B1) | B1 ile aynı dolgu rengine sahip hücreleri sayar. Biçim değişimi sonrası Ctrl+Alt+F9 kullanın.

Yardım ve Hakkında / HAMU_YerelDosyaYolu
=HAMU_YerelDosyaYolu(A1) | Kişisel OneDrive URL yolunu yerel OneDrive ortam değişkeniyle eşler; diğer URL türlerini korur.

Yardım ve Hakkında / HAMU_GizliLink
=HAMU_GizliLink(A1) | Hücrenin ilk köprü adresini getirir.

Yardım ve Hakkında / HAMU_LinkParamDegeri
=HAMU_LinkParamDegeri(A1;"id") | URL içindeki sorgu parametresinin değerini getirir.

Yardım ve Hakkında / HAMU_SayidanMetine
=HAMU_SayidanMetine(123) | Türkçe: yüz yirmi üç. 15 basamaktan küçük mutlak değerler; ondalık kısmı tek tek okur.

Yardım ve Hakkında / HAMU_TablodanVeriGetir
=HAMU_TablodanVeriGetir(A1;"Tablo1";"Kod";"Ad") | Formülün bulunduğu kitapta tablodan ilk eşleşen değeri getirir.

Kontrol ve Denetim / Renge Göre Say
Seçilen aralıkta örnek hücrenin görüntülenen dolgu rengiyle eşleşen hücreleri sayar; koşullu biçimlendirme dahildir.

Kontrol ve Denetim / Renge Göre Topla
Örnek hücrenin görüntülenen dolgu rengiyle eşleşen sayısal değerleri toplar; metin, boş ve hata değerlerini toplama katmaz.

Metin ve Hücre / Boşlukları Sıfırla Doldur
Seçili aralıktaki gerçekten boş hücrelere 0 yazar. Dolu hücreleri ve boş metin döndüren formülleri korur.

Metin ve Hücre / Başa / Sona Ekle
Seçili tek kolondaki metinlerin başına ve/veya sonuna metin ekler. Sonuç seçtiğiniz başlangıç hücresine yazılır; kaynak korunur. Boş hücreler boş kalır.

Metin ve Hücre / Baştaki Sıfırlar
Kodları sabit uzunluğa baştan sıfır ekleyerek düzenler veya baştaki sıfırları kaldırır. Sonuç metin olarak seçtiğiniz hedefe yazılır; uzun kodlar ve kaynak korunur.

Metin ve Hücre / Hücre İçi Tekrarlar
Virgül veya belirttiğiniz ayraçla ayrılmış hücre içi öğeleri tekilleştirir. Örneğin elma, armut, elma → elma, armut. İlk sıra korunur; sonuç ayrı hedefe yazılır.

Metin ve Hücre / Metin Parçası Al
Soldan/sağdan belirli sayıda karakteri, bir işaretin öncesini/sonrasını veya iki işaret arasını ayıklar. İlk eşleşmeyi kullanır; eşleşme yoksa boş sonuç verir. Kaynak korunur.

Tarih Araçları / Tarih Listesi
Başlangıç ve bitiş dahil tarih listesi oluşturur. Tüm günler veya yalnız hafta içi seçilir; hafta içi seçeneği resmî tatilleri hesaplamaz. Sonuç seçtiğiniz başlangıç hücresine yazılır.

Kontrol ve Denetim / Formül Denetimi
Kitaptaki formülleri kaynak hücre bağlantısı, formül metni, görüntülenen sonuç, hata durumu ve dış kitap başvurusu bilgisiyle yeni sayfada listeler. Kaynak formüller değiştirilmez. Dış başvuru tespiti doğrudan XLS/CSV kitap başvurularını tarar.

Kontrol ve Denetim / Özel Hücreleri Seç
Mevcut seçimdeki formül, sabit değer, hata, boş veya görünür hücreleri seçer. Değer ve biçimler değiştirilmez; boş metin döndüren formüller gerçek boş sayılmaz.

Kontrol ve Denetim / Tanımlı Adları Listele
Kitap ve sayfa kapsamındaki tanımlı adları, başvurularını, görünürlüklerini ve geçersiz başvuru durumlarını yeni sayfada listeler. Adları değiştirmez veya silmez.

## Ayarlar ve dil

Dosya > HAMU Tools > Ayarlar bölümünde yedek/PDF klasörleri ve ortak seçim/çıktı tercihleri düzenlenir. Dil seçimi Türkçe veya English olarak kalıcı saklanır. Ayrıntılı tercih tablosu için V1_Ayarlar_Dil_ve_Guvenilirlik.md dosyasını okuyun.
