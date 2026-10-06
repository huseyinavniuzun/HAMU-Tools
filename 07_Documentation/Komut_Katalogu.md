# HAMU Tools V1 — Güncel Komut Kataloğu

93 menü kaydı

## Aktarma Sihirbazı

Şablondan LAMBDA, hücre stili, adlandırılmış veri aralığı, sayfa verisi ve Excel tablosu aktarır. Sayfa/tablo aktarımı değer ve biçim kopyalar; VBA taşımaz.

Kod: `HAMU_Transfer`

## Yedek Oluştur

Dosyanın tarih ve saatli kopyasını Belgeler > HAMU > Backup klasörüne kaydeder. Sonuç ekranında tam yolu gösterir; Klasörü Aç ile yedek klasörüne ulaşabilirsiniz.

Kod: `HAMU_Backup`

## Tarihli Sürüm

Kaydedilmiş dosyanın aynı klasörde tarih ve saat eklenmiş kopyasını oluşturur. Yeni kopya açılıp doğrulandıktan sonra eski dosya kapatılır ve silinir. Yeni dosya açılamazsa eski dosya korunur.

Kod: `HAMU_AddDateToFilename`

## PDF Kaydet

Alan Seç veya Etkin Sayfa seçeneğiyle PDF kaydeder. Alan Seç her zaman Excel aralık seçicisini açar; mevcut seçimi kendiliğinden kaydetmez. Otomatik/dikey/yatay yön sunar.

Kod: `HAMU_ExportSheetsAsPDF`

## Klasör İçeriği

Klasör seçildikten sonra yalnız bu klasör veya alt klasörlerle birlikte tarama seçilir. Alt klasör taraması uzun sürebilir. Dosyaları bağlantı, yol, tür, okunabilir boyut ve tarihlerle listeler.

Kod: `HAMU_FolderIndex`

## CSV Dosyalarını Birleştir

Seçtiğiniz klasördeki CSV dosyalarını tek sayfada alt alta toplar; ilk dosyanın başlığını kullanır ve sonraki dosyalarda başlık satırını atlar.

Kod: `HAMU_CombineCSVs`

## Sayfaları Birleştir

Aktif çalışma kitabındaki sayfaların kullanılan alanlarını tek sayfada birleştirir ve kaynak sayfa bilgisini sonuç verisine ekler.

Kod: `HAMU_ConsolidateSheets`

## Tabloları Birleştir

Belirttiğiniz Excel tablolarının aynı yapıdaki satırlarını tek bir sonuç tablosunda alt alta toplar.

Kod: `HAMU_AppendSelectedTables`

## Akıllı Birleştir

Aynı bilgileri içeren sayfaları alt alta birleştirir. Örneğin Ocak ve Şubat satışlarında sütun sırası farklı olsa da Ürün başlıkları eşleştirilir. Eksik sütunlar boş kalır, KaynakSayfa eklenir. Kaynak sayfalar değiştirilmez; ilk satırlar başlık olmalıdır.

Kod: `HAMU_SmartAppendByHeaders`

## Dosyalardan Topla

Klasördeki Excel ve CSV dosyalarının ilk veri sayfalarını tek sonuçta alt alta toplar. İlk satır başlıktır; aynı başlıklar eşleştirilir. KaynakDosya eklenir, kaynak dosyalar değiştirilmez. Aynı tür tabloları içeren dosyalar kullanın.

Kod: `HAMU_CollectFilesData`

## Anahtara Göre Tablo Eşleştir

İki tablodaki kayıtları müşteri kodu veya ürün kodu gibi ortak alanlarla yan yana eşleştirir. LEFT: ilk tablonun tüm satırları; INNER: yalnız eşleşenler; FULL: iki tablonun tüm kayıtları. İkinci tablodan alınacak sütunlar ayrıca seçilir. Kaynaklar değiştirilmez.

Kod: `HAMU_SmartJoin`

## Grupla ve Özetle

Bir veya birden fazla kolona göre kayıtları gruplar; birden fazla sayısal kolonu aynı anda toplar. Kolonları numarayla veya başlık adıyla seçebilirsiniz.

Kod: `HAMU_GroupAndSummarize`

## Unpivot

Solda belirlediğiniz birden fazla sabit kolonu korur; sağa doğru uzanan pivot kolonlarını Alan ve Değer yapısında satırlara dönüştürür.

Kod: `HAMU_UnpivotColumnsToRows`

## Veriyi Böl

Seçili tabloyu satır sayısına, eşit parçalara, tek veya çoklu kolon değerlerine göre ya da yıl, çeyrek, ay ve hafta bazında tarih grupları oluşturarak böler. Çıktıları yeni çalışma sayfalarına veya ayrı Excel dosyalarına aktarabilirsiniz.

Kod: `HAMU_SplitDataWizard`

## Dosyalara Böl

Seçtiğiniz kolonun benzersiz değerlerine göre veriyi filtreler ve her değer için ayrı bir Excel dosyası oluşturur.

Kod: `HAMU_SplitByValueToFiles`

## URL Ayıkla

URL'lerde soru işaretinden sonra bulunan sorgu parametrelerini algılar ve her parametreyi ayrı bir kolon halinde yeni sayfaya çıkarır.

Kod: `HAMU_SplitURLParameters`

## Pivotu Böl

Seçili PivotTable'ın rapor filtresindeki değerleri sırayla uygular ve her filtre sonucu için ayrı çalışma sayfası oluşturur.

Kod: `HAMU_PivotExport`

## Pivot

Uzun formattaki veriyi tekrar geniş çapraz tabloya dönüştürür. Birden fazla sabit anahtar kolonu kullanabilir; yeni kolon başlıklarını seçilen alandan üretir ve aynı hücreye düşen kayıtları toplama, ilk değer veya sayım yöntemiyle birleştirir.

Kod: `HAMU_PivotRowsToColumns`

## Boşlukları Aşağı Doldur

Seçili aralıktaki boş hücreleri aynı kolondaki en yakın üst dolu değerle doldurur. Pivot veya rapor çıktılarındaki grup etiketlerini satırlara yaymak için kullanışlıdır.

Kod: `HAMU_FillDownBlanks`

## Tekrarları Birleştir

Tek veya çoklu anahtar kolonlarla tekrar eden kayıtları tek satırda birleştirir. Seçtiğiniz sayısal kolonları toplar, diğer alanlarda ilk dolu değeri korur ve kaç kaydın birleştiğini ayrıca yazar.

Kod: `HAMU_MergeDuplicates`

## Başlıkları Düzelt

Kolon başlıklarını teknik snake_case biçimine dönüştürebilir veya gereksiz boşlukları temizleyerek okunabilir başlık yapısını koruyabilir. Aynı başlık oluşursa otomatik olarak benzersiz ad üretir.

Kod: `HAMU_StandardizeHeaders`

## Veri Türünü Düzelt

Bir veya birden fazla kolonu metin, sayı, tarih, tam sayı veya yüzde tipine dönüştürür; dönüştürülemeyen hücreleri sayarak işlem sonucunu raporlar.

Kod: `HAMU_ConvertColumnTypes`

## Kolon Seç

Başlık ve veri alanından kolonları adlarıyla veya seçim içindeki 1,2,3 gibi numaralarla seçer. Tüm sütun seçimi dolu son satıra göre daraltılır.

Kod: `HAMU_SelectColumnsToNewTable`

## Koşulla Satır Filtrele

Koşula uyan satırları yeni sayfaya kopyalar; kaynak değişmez. Şehir=Bursa yalnız Bursa kayıtlarını, Satış>0 pozitif satışları, Ürün~TV içinde TV geçenleri seçer. Noktalı virgülle birleştirilen koşulların hepsi sağlanmalıdır.

Kod: `HAMU_FilterRowsToNewTable`

## Veriden Örnek Al

Tüm veri kümesinden rastgele N satır seçebilir veya belirlediğiniz grup kolonlarına göre her gruptan en fazla N rastgele kayıt alarak yeni bir örnek veri sayfası oluşturabilir.

Kod: `HAMU_SampleData`

## Satırları Karıştır

Kaynak veriyi değiştirmeden, başlık satırını koruyarak veri satırlarını rastgele sırada yeni bir çalışma sayfasına aktarır.

Kod: `HAMU_ShuffleRows`

## İki Listeyi Karşılaştır

İki tek kolonlu listeyi karşılaştırır; yalnızca birinci listede, yalnızca ikinci listede veya her iki listede bulunan değerleri raporlar.

Kod: `HAMU_CompareTwoLists`

## Benzersiz Liste

Tekrar eden değerleri bir kez listeler. Çıktı için mevcut sayfada veya başka sayfada başlangıç hücresi seçebilir ya da yeni sayfa oluşturabilirsiniz. Dolu hedefte onay ister; kaynak hata ve boş hücrelerini atlar.

Kod: `HAMU_ExtractUniqueList`

## Veri Profili

Seçtiğiniz başlıklı veri alanındaki sütunların doluluk, boşluk, benzersiz değer ve veri türlerini özetler. Sorunları görmek için bir rapor oluşturur; kaynak veriyi değiştirmez.

Kod: `HAMU_DataProfile`

## Kalite Raporu

Başlıklı tabloyu denetleyerek boş alan, tekrar eden kayıt ve veri türü sorunlarını raporlar. Sonuç yeni sayfaya yazılır; sorunları kaynaktan otomatik silmez. Önce raporu inceleyin.

Kod: `HAMU_DataQualityReport`

## Eksikleri Bul

Mağaza, Ürün, Hafta gibi seçtiğiniz kolonların benzersiz değerlerinden tüm olası kombinasyonları oluşturur ve kaynak veri setinde bulunmayan kombinasyonları yeni bir sayfada listeler.

Kod: `HAMU_FindMissingCombinations`

## Kırp ve Temizle

Seçili metinlerde görünmeyen NBSP karakterlerini, sekmeleri, satır sonlarını, sıfır genişlikli boşlukları ve art arda gelen gereksiz boşlukları temizler. Metnin başı/sonundaki boşlukları kaldırır; formülleri değiştirmez.

Kod: `HAMU_CleanData`

## Harf Dönüştür

Seçili metinleri tamamı büyük, tamamı küçük veya baş harfleri büyük biçime dönüştürür.

Kod: `HAMU_TextCase`

## Normal İfade ile Bul

Seçili hücrelerde RegEx normal ifade desenine uyan metinleri bulur ve eşleşen hücreleri vurgular. Sayı, kod, telefon veya belirli metin kalıplarını tespit etmek için kullanılabilir.

Kod: `HAMU_PatternFind`

## Normal İfade ile Değiştir

Seçili metin hücrelerinde RegEx normal ifadeleri kullanarak toplu bul/değiştir işlemi yapar. Gruplama ve $1 gibi geri başvurular kullanılabilir.

Kod: `HAMU_RegexReplace`

## Metinden Sayıları Ayıkla

Metin içindeki rakam karakterlerini ayıklar ve sonucu hedef alana aktarır.

Kod: `HAMU_ExtractDigits`

## Metin Olarak Saklanan Sayıları Düzelt

Başında gizli tek tırnak bulunan veya Excel tarafından metin olarak saklanan sayısal değerleri gerçek sayıya dönüştürmeye çalışır.

Kod: `HAMU_ConvertQuotedNumbers`

## Veriyi Maskele

Seçili metinlerin belirlediğiniz kadar başlangıç ve bitiş karakterini koruyup ortadaki karakterleri yıldızla gizler.

Kod: `HAMU_MaskData`

## Formülleri Değere Dönüştür

Seçili hücrelerdeki formülleri mevcut sonuçlarıyla değiştirir. İşlem sonrasında hücrelerde formül yerine sabit değer kalır.

Kod: `HAMU_ConvertFormulasToValues`

## Birleşimleri Çöz ve Doldur

Birleştirilmiş hücreleri çözer ve birleşik alanın sol üst değerini çözülen hücrelerin tamamına yazar.

Kod: `HAMU_UnmergeAndFill`

## Standart Sayı ve Tarih Biçimi

Gerçek sayıları ondalıksız tam sayı görünümüyle, tarih hücrelerini gg.aa.yyyy biçiminde gösterir. Hücre değerlerini yuvarlamaz veya değiştirmez.

Kod: `HAMU_StandardFormat`

## Boşlukları Sil

Seçili alandaki boş hücreleri silerek alttaki hücreleri yukarı kaydırır. Veri hizasını etkileyebileceği için tek kolonlu listelerde kullanılması önerilir.

Kod: `HAMU_DeleteBlankCells`

## Eski Dropdown Nesnelerini Sil

Aktif sayfadaki eski Form Denetimi ve ActiveX açılır kutu nesnelerini siler. Hücrelerdeki veri doğrulama listeleri korunur.

Kod: `HAMU_DeleteDropdowns`

## Satırları Renklendir

Seçili aralıkta dönüşümlü satır renklendirmesi uygular. Çalıştırıldığında mavi, yeşil, sarı, turuncu, mor, gri ve pembe tonlarından birini seçebilirsiniz.

Kod: `HAMU_RowStriping`

## Tarih Seçici

HAMU takvim penceresini açar. Aktif hücrede tarih varsa takvim o tarihle açılır; seçtiğiniz tarih hücreye gerçek Excel tarihi olarak dd.mm.yyyy biçiminde yazılır.

Kod: `HAMU_DatePicker`

## Tarihleri Düzelt

03/10/2026, 2026-10-03, 20261003, 3 Ekim 2026 ve benzeri farklı tarih gösterimlerini gerçek Excel tarihine dönüştürür ve dd.mm.yyyy biçiminde standardize eder.

Kod: `HAMU_FixDates`

## Tarih Kontrolü

Seçili hücrelerdeki değerlerin geçerli bir tarih olup olmadığını kontrol eder. Geçerli tarihleri ve tanımlanamayan değerleri farklı renklerle işaretler.

Kod: `HAMU_CheckDates`

## Tarih Oluştur

Ayrı kolonlarda bulunan yıl, ay ve gün değerlerini birleştirerek gerçek Excel tarihleri oluşturur. Ay alanında sayı veya Türkçe ay adı kullanılabilir.

Kod: `HAMU_DateBuilder`

## Sayfa İndeksi

Görünür çalışma sayfalarını listeleyen ve her sayfaya tıklanabilir bağlantı veren bir İçindekiler sayfası oluşturur.

Kod: `HAMU_SheetIndex`

## Sayfaları Alfabetik Sırala

Çalışma kitabındaki çalışma sayfalarını adlarına göre alfabetik olarak yeniden sıralar.

Kod: `HAMU_SortSheets`

## Tüm Gizlileri Göster

Aktif çalışma kitabındaki gizli çalışma sayfalarını, satırları ve sütunları görünür hale getirir.

Kod: `HAMU_ShowHidden`

## Dinamik Veri Doğrulama Listesi

Seçili hücrelere liste tipi veri doğrulama uygular. Kaynak olarak adlandırılmış aralık, Tablo[Kolon], doğrudan aralık veya A1# biçimindeki dinamik taşan dizi kullanılabilir. Bir adlandırma A1# kaynağına bağlıysa taşan aralık büyüdükçe doğrulama listesi de dinamik kalır.

Kod: `HAMU_DynamicValidation`

## Resim Kaydet

Seçili alanı PNG, JPG veya GIF olarak kaydeder. Çok büyük görsel alanlarını durdurur ve geçici grafiği temizler.

Kod: `HAMU_SelectionToPNG`

## Toplu Fotoğraf

Seçtiğiniz klasördeki görselleri aktif hücreden başlayarak satırlara ekler ve yan hücreye dosya adını yazar.

Kod: `HAMU_InsertPhotosFromFolder`

## Karekod ve Barkod

Tek ekranda Karekod veya Barkod seçin; tür değişince model listesi güncellenir. Metin, bağlantı, e-posta ve telefon içeriği; tek giriş veya tek kolondan toplu üretim. Kodlar verinin sağındaki boş hücreye eklenir. İnternet gerekir; en fazla 250 kod.

Kod: `HAMU_CreateQRCode`

## Sayfaları Karşılaştır

İki çalışma sayfasını belirlediğiniz anahtar kolona göre karşılaştırır ve eklenen, silinen veya değişen kayıtları ayrı bir raporda gösterir.

Kod: `HAMU_CompareSheets`

## Sorunlu Hücreler

Sayı gibi görünen metinleri ve eşittir işaretiyle başlayan ancak gerçek formül olmayan metinleri bularak hücreleri farklı renklerle işaretler.

Kod: `HAMU_MarkProblemCells`

## Hücre Türüne Göre Vurgula

Seçili alanda sayı, metin, hata veya boş hücreleri seçtiğiniz kritere göre vurgular.

Kod: `HAMU_HighlightByCriteria`

## Tekrar Edenleri Vurgula

Seçili alanda birden fazla kez bulunan değerleri tespit eder ve tekrar eden hücreleri renklendirir.

Kod: `HAMU_HighlightDuplicates`

## Tekrar Sayısını Yaz

Giriş listesindeki her değerin kaç kez geçtiğini hesaplayıp ayrı seçtiğiniz çıkış alanına yazar. Tek çıkış hücresi seçerseniz sonuç aşağı doğru genişler.

Kod: `HAMU_AddDuplicateCount`

## Harici Bağlantıları Listele

Çalışma kitabındaki Excel türü dış bağlantıları yeni bir sayfada listeler.

Kod: `HAMU_ListLinks`

## Harici Bağlantıları Kır

Çalışma kitabındaki harici Excel bağlantılarını kırar. Bağlantılı formüller mevcut değerlerine dönüştürülebileceği için işlemden önce yedek alınması önerilir.

Kod: `HAMU_BreakLinks`

## Yardım

HAMU Tools içindeki komutların kullanım amacı ve temel kullanım bilgilerini gösterir.

Kod: `HAMU_Help`

## Hakkında

HAMU Tools V1 sürümü, HAMU markası ve huseyinavniuzun.com bilgileri ile eklentinin kapsamını gösterir.

Kod: `HAMU_About`

## Sığdır

Sütun genişliğini ve satır yüksekliğini içeriğe göre ayarlar; tabloya yeni satır eklemez.

Kod: `HAMU_AutoFit`

## Otomatik Tablo Oluştur

Tek hücre seçildiyse çevresindeki kesintisiz veri alanını, aksi halde seçilen alanı başlıklı Excel tablosuna dönüştürür. HAMU stili ve filtreleri uygular.

Kod: `HAMU_AutoTable`

## Tabloyu Veriye Genişlet

Seçili tablonun altındaki kesintisiz veriyi aynı sütun genişliğinde tabloya dahil eder. Başka tablolarla çakışmayı önler.

Kod: `HAMU_ResizeTable`

## Çalışma Sayfası Fonksiyonları

HAMU çalışma sayfası fonksiyonlarını ve formül örneklerini gösterir.

Kod: `HAMU_UDFHelp`

## HAMU_URLCoz

=HAMU_URLCoz(A1) | UTF-8 URL karakterlerini çözer. Geçersiz yüzde dizisini korur.

Kod: `UDF_URLDecode`

## HAMU_RengeGoreTopla

=HAMU_RengeGoreTopla(A1:A20;B1) | B1 ile aynı dolgu rengine sahip sayıları toplar. Biçim değişimi sonrası F9 kullanın.

Kod: `UDF_RenklileriTopla`

## HAMU_RengeGoreSay

=HAMU_RengeGoreSay(A1:A20;B1) | B1 ile aynı dolgu rengine sahip hücreleri sayar. Biçim değişimi sonrası F9 kullanın.

Kod: `UDF_RengeGoreSay`

## HAMU_YerelDosyaYolu

=HAMU_YerelDosyaYolu(A1) | Kişisel OneDrive URL yolunu yerel OneDrive ortam değişkeniyle eşler; diğer URL türlerini korur.

Kod: `UDF_YerelDosyaYolu`

## HAMU_GizliLink

=HAMU_GizliLink(A1) | Hücrenin ilk köprü adresini getirir.

Kod: `UDF_GizliLink`

## HAMU_LinkParamDegeri

=HAMU_LinkParamDegeri(A1;"id") | URL içindeki sorgu parametresinin değerini getirir.

Kod: `UDF_LinkParamDegeri`

## HAMU_SayidanMetine

=HAMU_SayidanMetine(123) | Türkçe: yüz yirmi üç. 15 basamaktan küçük mutlak değerler; ondalık kısmı tek tek okur.

Kod: `UDF_SayidanMetine`

## HAMU_TablodanVeriGetir

=HAMU_TablodanVeriGetir(A1;"Tablo1";"Kod";"Ad") | Formülün bulunduğu kitapta tablodan ilk eşleşen değeri getirir.

Kod: `UDF_TablodanVeriGetir`

## Renge Göre Say

Seçilen aralıkta örnek hücrenin görüntülenen dolgu rengiyle eşleşen hücreleri sayar; koşullu biçimlendirme dahildir.

Kod: `HAMU_CountByColor`

## Renge Göre Topla

Örnek hücrenin görüntülenen dolgu rengiyle eşleşen sayısal değerleri toplar; metin, boş ve hata değerlerini toplama katmaz.

Kod: `HAMU_SumByColor`

## Boşlukları Sıfırla Doldur

Seçili aralıktaki gerçekten boş hücrelere 0 yazar. Dolu hücreleri ve boş metin döndüren formülleri korur.

Kod: `HAMU_FillBlanksZero`

## Başa / Sona Ekle

Seçili tek kolondaki metinlerin başına ve/veya sonuna metin ekler. Sonuç seçtiğiniz başlangıç hücresine yazılır; kaynak korunur. Boş hücreler boş kalır.

Kod: `HAMU_AddTextAffixes`

## Baştaki Sıfırlar

Kodları sabit uzunluğa baştan sıfır ekleyerek düzenler veya baştaki sıfırları kaldırır. Sonuç metin olarak seçtiğiniz hedefe yazılır; uzun kodlar ve kaynak korunur.

Kod: `HAMU_LeadingZeros`

## Hücre İçi Tekrarlar

Virgül veya belirttiğiniz ayraçla ayrılmış hücre içi öğeleri tekilleştirir. Örneğin elma, armut, elma → elma, armut. İlk sıra korunur; sonuç ayrı hedefe yazılır.

Kod: `HAMU_DeduplicateItems`

## Metin Parçası Al

Soldan/sağdan belirli sayıda karakteri, bir işaretin öncesini/sonrasını veya iki işaret arasını ayıklar. İlk eşleşmeyi kullanır; eşleşme yoksa boş sonuç verir. Kaynak korunur.

Kod: `HAMU_ExtractTextPart`

## Tarih Listesi

Başlangıç ve bitiş dahil tarih listesi oluşturur. Tüm günler veya yalnız hafta içi seçilir; hafta içi seçeneği resmî tatilleri hesaplamaz. Sonuç seçtiğiniz başlangıç hücresine yazılır.

Kod: `HAMU_DateRangeList`

## Formül Denetimi

Kitaptaki formülleri kaynak hücre bağlantısı, formül metni, görüntülenen sonuç, hata durumu ve dış kitap başvurusu bilgisiyle yeni sayfada listeler. Kaynak formüller değiştirilmez. Dış başvuru tespiti doğrudan XLS/CSV kitap başvurularını tarar.

Kod: `HAMU_FormulaAudit`

## Özel Hücreleri Seç

Mevcut seçimdeki formül, sabit değer, hata, boş veya görünür hücreleri seçer. Değer ve biçimler değiştirilmez; boş metin döndüren formüller gerçek boş sayılmaz.

Kod: `HAMU_SelectSpecialCells`

## Tanımlı Adları Listele

Kitap ve sayfa kapsamındaki tanımlı adları, başvurularını, görünürlüklerini ve geçersiz başvuru durumlarını yeni sayfada listeler. Adları değiştirmez veya silmez.

Kod: `HAMU_ListDefinedNames`

## Veri Girişi

Excel’in yerleşik veri giriş formunu açar. Bir tabloya veya benzersiz başlıkları olan kesintisiz listeye tıklayın; en fazla 32 kolon.

Kod: `HAMU_DataEntry`

## Vurguyu Kaldır

Seçili alandaki HAMU vurgu renklerini kaldırır. Aynı oturumda önceki dolguyu geri yükler; daha eski vurgular sarı/turuncu paletiyle tanınır.

Kod: `HAMU_ClearHighlights`

## HAMU_L_CokluAra

İki ayrı anahtar/sonuç kolonunda tam eşleşme arar; ilk sonucu verir. Örnek: =HAMU_L_CokluAra(G2;A2:A10;B2:B10;D2:D10;E2:E10). LAMBDA şablonundan aktarılır; Microsoft 365/Excel 2024 gerekir.

Kod: `UDF_LambdaMultiLookup`

## HAMU_L_CokluTopla

İki ayrı anahtar/değer kolonunda koşula uyan sayıları toplar. Örnek: =HAMU_L_CokluTopla(G2;A2:A10;B2:B10;D2:D10;E2:E10). Şablondan aktarılır; Microsoft 365/Excel 2024 gerekir.

Kod: `UDF_LambdaMultiSum`

## HAMU_L_BenzersizBirlestir

İki alandaki dolu değerleri tek bir sıralı benzersiz listeye dönüştürür. Örnek: =HAMU_L_BenzersizBirlestir(A2:B10;D2:E10). Sonuç için boş taşma alanı gerekir. Microsoft 365/Excel 2024 gerekir.

Kod: `UDF_LambdaUniqueMerge`

## Kopya Kaydet

Dosyanın bir kopyasını seçtiğiniz konuma kaydeder. Açık dosyanın adı, konumu ve kaydedilmemiş değişiklik durumu korunur; biçim ve makrolar değiştirilmez.

Kod: `HAMU_SaveCopy`
