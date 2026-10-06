Attribute VB_Name = "modHAMU_CommandInfo"
Option Explicit
Option Private Module
Public gHAMU_CommandId As String
Public Function HAMU_CommandIntro(ByVal id As String) As String
 Select Case id
 Case "HAMU_Transfer"
 HAMU_CommandIntro = HAMU_L("Şablondan LAMBDA, hücre stili, adlandırılmış veri aralığı, sayfa verisi ve Excel tablosu aktarır. Sayfa/tablo aktarımı değer ve biçim kopyalar; VBA taşımaz.", "Transfer LAMBDA definitions, cell styles, named data ranges, sheet data and Excel tables. Sheet/table transfer copies values and formatting, without VBA.")
 Case "HAMU_Backup"
 HAMU_CommandIntro = HAMU_L("Dosyanın tarih ve saatli kopyasını Belgeler > HAMU > Backup klasörüne kaydeder. Sonuç ekranında tam yolu gösterir; Klasörü Aç ile yedek klasörüne ulaşabilirsiniz.", "Save a timestamped copy to the backup folder configured in Settings. The result displays its full path and an Open Folder button.")
 Case "HAMU_AddDateToFilename"
 HAMU_CommandIntro = HAMU_L("Kaydedilmiş dosyanın aynı klasörde tarih ve saat eklenmiş kopyasını oluşturur. Yeni kopya açılıp doğrulandıktan sonra eski dosya kapatılır ve silinir. Yeni dosya açılamazsa eski dosya korunur.", "Create a timestamped copy in the original folder. Open and verify the copy before closing and deleting the original; keep the original if verification fails.")
 Case "HAMU_ExportSheetsAsPDF"
 HAMU_CommandIntro = HAMU_L("Alan Seç veya Etkin Sayfa seçeneğiyle PDF kaydeder. Alan Seç her zaman Excel aralık seçicisini açar; mevcut seçimi kendiliğinden kaydetmez. Otomatik/dikey/yatay yön sunar.", "Save PDF using Select Range or Active Sheet. Select Range always opens Excel's range picker instead of automatically exporting the selection. Choose automatic, portrait or landscape.")
 Case "HAMU_FolderIndex"
 HAMU_CommandIntro = HAMU_L("Klasör seçildikten sonra yalnız bu klasör veya alt klasörlerle birlikte tarama seçilir. Alt klasör taraması uzun sürebilir. Dosyaları bağlantı, yol, tür, okunabilir boyut ve tarihlerle listeler.", "List files with hyperlinks, local paths, extensions, types, readable sizes and creation/modification dates. Choose current-folder or recursive scanning; large folder trees take longer.")
 Case "HAMU_CombineCSVs"
 HAMU_CommandIntro = HAMU_L("Seçtiğiniz klasördeki CSV dosyalarını tek sayfada alt alta toplar; ilk dosyanın başlığını kullanır ve sonraki dosyalarda başlık satırını atlar.", "Append CSV files from a folder to one sheet. Keep the first file's header and skip header rows in subsequent files.")
 Case "HAMU_ConsolidateSheets"
 HAMU_CommandIntro = HAMU_L("Aktif çalışma kitabındaki sayfaların kullanılan alanlarını tek sayfada birleştirir ve kaynak sayfa bilgisini sonuç verisine ekler.", "Append used ranges from worksheets in the active workbook and add the source worksheet name to the results.")
 Case "HAMU_AppendSelectedTables"
 HAMU_CommandIntro = HAMU_L("Belirttiğiniz Excel tablolarının aynı yapıdaki satırlarını tek bir sonuç tablosunda alt alta toplar.", "Combine rows from selected Excel tables with matching structures into one result table.")
 Case "HAMU_SmartAppendByHeaders"
 HAMU_CommandIntro = HAMU_L("Aynı bilgileri içeren sayfaları alt alta birleştirir. Örneğin Ocak ve Şubat satışlarında sütun sırası farklı olsa da Ürün başlıkları eşleştirilir. Eksik sütunlar boş kalır, KaynakSayfa eklenir. Kaynak sayfalar değiştirilmez; ilk satırlar başlık olmalıdır.", "Append worksheets by matching header names, even when column order differs. Missing columns remain blank. Add the source sheet name and preserve the source worksheets.")
 Case "HAMU_CollectFilesData"
 HAMU_CommandIntro = HAMU_L("Klasördeki Excel ve CSV dosyalarının ilk veri sayfalarını tek sonuçta alt alta toplar. İlk satır başlıktır; aynı başlıklar eşleştirilir. KaynakDosya eklenir, kaynak dosyalar değiştirilmez. Aynı tür tabloları içeren dosyalar kullanın.", "Read the first data sheet from Excel/CSV files in a folder and append matching headers. Add the source filename without changing source files.")
 Case "HAMU_SmartJoin"
 HAMU_CommandIntro = HAMU_L("İki tablodaki kayıtları müşteri kodu veya ürün kodu gibi ortak alanlarla yan yana eşleştirir. LEFT: ilk tablonun tüm satırları; INNER: yalnız eşleşenler; FULL: iki tablonun tüm kayıtları. İkinci tablodan alınacak sütunlar ayrıca seçilir. Kaynaklar değiştirilmez.", "Match two datasets using selected key columns. Check key types, spaces and duplicates before choosing the fields to include.")
 Case "HAMU_GroupAndSummarize"
 HAMU_CommandIntro = HAMU_L("Bir veya birden fazla sütuna göre kayıtları gruplar; birden fazla sayısal sütunu aynı anda toplar. Sütunları numarayla veya başlık adıyla seçebilirsiniz.", "Group records by one or more key columns and summarize selected numeric fields. Specify column numbers or header names.")
 Case "HAMU_UnpivotColumnsToRows"
 HAMU_CommandIntro = HAMU_L("Solda belirlediğiniz birden fazla sabit sütunu korur; sağa doğru uzanan pivot sütunlarını Alan ve Değer yapısında satırlara dönüştürür.", "Convert wide columns into field/value rows. Keep text headers such as 2026.01 as text.")
 Case "HAMU_SplitDataWizard"
 HAMU_CommandIntro = HAMU_L("Seçili tabloyu satır sayısına, eşit parçalara, tek veya çoklu sütun değerlerine göre ya da yıl, çeyrek, ay ve hafta bazında tarih grupları oluşturarak böler. Çıktıları yeni çalışma sayfalarına veya ayrı Excel dosyalarına aktarabilirsiniz.", "Split values or datasets using the options shown in the dialog. Choose the source columns and desired output.")
 Case "HAMU_SplitByValueToFiles"
 HAMU_CommandIntro = HAMU_L("Seçtiğiniz sütunun benzersiz değerlerine göre veriyi filtreler ve her değer için ayrı bir Excel dosyası oluşturur.", "Create separate files for distinct values in a selected column. Choose the destination folder and preserve the source data.")
 Case "HAMU_SplitURLParameters"
 HAMU_CommandIntro = HAMU_L("URL'lerde soru işaretinden sonra bulunan sorgu parametrelerini algılar ve her parametreyi ayrı bir sütun halinde yeni sayfaya çıkarır.", "Extract URL query parameters into separate columns for analysis.")
 Case "HAMU_PivotExport"
 HAMU_CommandIntro = HAMU_L("Seçili PivotTable'ın rapor filtresindeki değerleri sırayla uygular ve her filtre sonucu için ayrı çalışma sayfası oluşturur.", "Export separate PivotTable filter results to worksheets. Select a cell within the PivotTable first.")
 Case "HAMU_PivotRowsToColumns"
 HAMU_CommandIntro = HAMU_L("Uzun formattaki veriyi tekrar geniş çapraz tabloya dönüştürür. Birden fazla sabit anahtar sütunu kullanabilir; yeni sütun başlıklarını seçilen alandan üretir ve aynı hücreye düşen kayıtları toplama, ilk değer veya sayım yöntemiyle birleştirir.", "Build a cross-tab by moving row categories into columns and summarizing values.")
 Case "HAMU_FillDownBlanks"
 HAMU_CommandIntro = HAMU_L("Seçili aralıktaki boş hücreleri aynı sütundaki en yakın üst dolu değerle doldurur. Pivot veya rapor çıktılarındaki grup etiketlerini satırlara yaymak için kullanışlıdır.", "Fill blank cells with the nearest nonblank value above in the same column. Useful for repeated group labels in reports.")
 Case "HAMU_MergeDuplicates"
 HAMU_CommandIntro = HAMU_L("Tek veya çoklu anahtar sütunlarla tekrar eden kayıtları tek satırda birleştirir. Seçtiğiniz sayısal sütunları toplar, diğer alanlarda ilk dolu değeri korur ve kaç kaydın birleştiğini ayrıca yazar.", "Merge records by one or more key columns, sum selected numeric fields, keep the first nonblank value in other fields and include a record count.")
 Case "HAMU_StandardizeHeaders"
 HAMU_CommandIntro = HAMU_L("Sütun başlıklarını teknik snake_case biçimine dönüştürebilir veya gereksiz boşlukları temizleyerek okunabilir başlık yapısını koruyabilir. Aynı başlık oluşursa otomatik olarak benzersiz ad üretir.", "Clean spaces and normalize column headings to a consistent naming scheme.")
 Case "HAMU_ConvertColumnTypes"
 HAMU_CommandIntro = HAMU_L("Bir veya birden fazla sütunu metin, sayı, tarih, tam sayı veya yüzde tipine dönüştürür; dönüştürülemeyen hücreleri sayarak işlem sonucunu raporlar.", "Convert selected columns to the requested data types. Review mixed values and conversion errors.")
 Case "HAMU_SelectColumnsToNewTable"
 HAMU_CommandIntro = HAMU_L("Başlık ve veri alanından sütunları adlarıyla veya seçim içindeki 1,2,3 gibi numaralarla seçer. Seçilen sütunların yeni bir kopyasını üretir.", "Select columns by header names or positions such as 1,2,3 within the selected data. Whole-column sources are trimmed to populated rows.")
 Case "HAMU_FilterRowsToNewTable"
 HAMU_CommandIntro = HAMU_L("Koşula uyan satırları yeni sayfaya kopyalar; kaynak değişmez. Şehir=Bursa yalnız Bursa kayıtlarını, Satış>0 pozitif satışları, Ürün~TV içinde TV geçenleri seçer. Noktalı virgülle birleştirilen koşulların hepsi sağlanmalıdır.", "Create an output table containing rows that satisfy the chosen column condition. Source data remains unchanged.")
 Case "HAMU_SampleData"
 HAMU_CommandIntro = HAMU_L("Tüm veri kümesinden rastgele N satır seçebilir veya belirlediğiniz grup sütunlarına göre her gruptan en fazla N rastgele kayıt alarak yeni bir örnek veri sayfası oluşturabilir.", "Create a sample of records using the requested sample size or selection options.")
 Case "HAMU_ShuffleRows"
 HAMU_CommandIntro = HAMU_L("Kaynak veriyi değiştirmeden, başlık satırını koruyarak veri satırlarını rastgele sırada yeni bir çalışma sayfasına aktarır.", "Randomize the order of data rows while preserving values within each row.")
 Case "HAMU_CompareTwoLists"
 HAMU_CommandIntro = HAMU_L("İki tek sütunlu listeyi karşılaştırır; yalnızca birinci listede, yalnızca ikinci listede veya her iki listede bulunan değerleri raporlar.", "Compare values in two lists and report matches and differences.")
 Case "HAMU_ExtractUniqueList"
 HAMU_CommandIntro = HAMU_L("Tekrar eden değerleri bir kez listeler. Çıktı için mevcut sayfada veya başka sayfada başlangıç hücresi seçebilir ya da yeni sayfa oluşturabilirsiniz. Dolu hedefte onay ister; kaynak hata ve boş hücrelerini atlar.", "Create a list of distinct values in a selected destination range or on a new worksheet. Configure the default destination behavior in Settings.")
 Case "HAMU_DataProfile"
 HAMU_CommandIntro = HAMU_L("Seçtiğiniz başlıklı veri alanındaki sütunların doluluk, boşluk, benzersiz değer ve veri türlerini özetler. Sorunları görmek için bir rapor oluşturur; kaynak veriyi değiştirmez.", "Summarize the contents and data types of selected columns to help assess a dataset.")
 Case "HAMU_DataQualityReport"
 HAMU_CommandIntro = HAMU_L("Başlıklı tabloyu denetleyerek boş alan, tekrar eden kayıt ve veri türü sorunlarını raporlar. Sonuç yeni sayfaya yazılır; sorunları kaynaktan otomatik silmez. Önce raporu inceleyin.", "Inspect missing values, duplicates and data consistency. Review the report before deciding how to clean the source data.")
 Case "HAMU_FindMissingCombinations"
 HAMU_CommandIntro = HAMU_L("Mağaza, Ürün, Hafta gibi seçtiğiniz sütunların benzersiz değerlerinden tüm olası kombinasyonları oluşturur ve kaynak veri setinde bulunmayan kombinasyonları yeni bir sayfada listeler.", "Find missing combinations of selected category columns.")
 Case "HAMU_CleanData"
 HAMU_CommandIntro = HAMU_L("Seçili metinlerde görünmeyen NBSP karakterlerini, sekmeleri, satır sonlarını, sıfır genişlikli boşlukları ve art arda gelen gereksiz boşlukları temizler. Metnin başı/sonundaki boşlukları kaldırır; formülleri değiştirmez.", "Clean redundant spaces and invisible characters while preserving Turkish characters.")
 Case "HAMU_TextCase"
 HAMU_CommandIntro = HAMU_L("Seçili metinleri tamamı büyük, tamamı küçük veya baş harfleri büyük biçime dönüştürür.", "Choose uppercase, lowercase or title case from a simple selection dialog.")
 Case "HAMU_PatternFind"
 HAMU_CommandIntro = HAMU_L("Seçili hücrelerde RegEx normal ifade desenine uyan metinleri bulur ve eşleşen hücreleri vurgular. Sayı, kod, telefon veya belirli metin kalıplarını tespit etmek için kullanılabilir.", "Find text matching a regular-expression pattern and highlight matching cells. Use a preset or a custom pattern.")
 Case "HAMU_RegexReplace"
 HAMU_CommandIntro = HAMU_L("Seçili metin hücrelerinde RegEx normal ifadeleri kullanarak toplu bul/değiştir işlemi yapar. Gruplama ve $1 gibi geri başvurular kullanılabilir.", "Replace matching text using a regular expression. Capture groups and references such as $1 can be used.")
 Case "HAMU_ExtractDigits"
 HAMU_CommandIntro = HAMU_L("Metin içindeki rakam karakterlerini ayıklar ve sonucu hedef alana aktarır.", "Extract numeric characters from selected text values.")
 Case "HAMU_ConvertQuotedNumbers"
 HAMU_CommandIntro = HAMU_L("Başında gizli tek tırnak bulunan veya Excel tarafından metin olarak saklanan sayısal değerleri gerçek sayıya dönüştürmeye çalışır.", "Convert suitable numbers stored as text to numeric values. Review decimal and thousands separators.")
 Case "HAMU_MaskData"
 HAMU_CommandIntro = HAMU_L("Seçili metinlerin belirlediğiniz kadar başlangıç ve bitiş karakterini koruyup ortadaki karakterleri yıldızla gizler.", "Mask selected text according to the available options. Keep an original copy when required.")
 Case "HAMU_ConvertFormulasToValues"
 HAMU_CommandIntro = HAMU_L("Seçili hücrelerdeki formülleri mevcut sonuçlarıyla değiştirir. İşlem sonrasında hücrelerde formül yerine sabit değer kalır.", "Replace selected formulas with their current calculated results.")
 Case "HAMU_UnmergeAndFill"
 HAMU_CommandIntro = HAMU_L("Birleştirilmiş hücreleri çözer ve birleşik alanın sol üst değerini çözülen hücrelerin tamamına yazar.", "Unmerge cells and fill the resulting area with the original merged value.")
 Case "HAMU_StandardFormat"
 HAMU_CommandIntro = HAMU_L("Gerçek sayıları ondalıksız tam sayı görünümüyle, tarih hücrelerini gg.aa.yyyy biçiminde gösterir. Hücre değerlerini yuvarlamaz veya değiştirmez.", "Apply integer number formatting or standard date formatting. Formatting controls display and does not automatically repair data types.")
 Case "HAMU_DeleteBlankCells"
 HAMU_CommandIntro = HAMU_L("Seçili alandaki boş hücreleri silerek alttaki hücreleri yukarı kaydırır. Veri hizasını etkileyebileceği için tek sütunlu listelerde kullanılması önerilir.", "Delete empty cells and shift remaining cells upward. Check related columns before changing their alignment.")
 Case "HAMU_DeleteDropdowns"
 HAMU_CommandIntro = HAMU_L("Aktif sayfadaki eski Form Denetimi ve ActiveX açılır kutu nesnelerini siler. Hücrelerdeki veri doğrulama listeleri korunur.", "Remove old worksheet dropdown objects while preserving cell data-validation lists.")
 Case "HAMU_RowStriping"
 HAMU_CommandIntro = HAMU_L("Seçili aralıkta dönüşümlü satır renklendirmesi uygular. Çalıştırıldığında mavi, yeşil, sarı, turuncu, mor, gri ve pembe tonlarından birini seçebilirsiniz.", "Apply alternating row colors to the selected area.")
 Case "HAMU_DatePicker"
 HAMU_CommandIntro = HAMU_L("HAMU takvim penceresini açar. Aktif hücrede tarih varsa takvim o tarihle açılır; seçtiğiniz tarih hücreye gerçek Excel tarihi olarak dd.mm.yyyy biçiminde yazılır.", "Choose a date and write it to the previously selected active cell. Select month/year and see the distance from today.")
 Case "HAMU_FixDates"
 HAMU_CommandIntro = HAMU_L("03/10/2026, 2026-10-03, 20261003, 3 Ekim 2026 ve benzeri farklı tarih gösterimlerini gerçek Excel tarihine dönüştürür ve dd.mm.yyyy biçiminde standardize eder.", "Normalize suitable date values and review values that cannot be converted.")
 Case "HAMU_CheckDates"
 HAMU_CommandIntro = HAMU_L("Seçili hücrelerdeki değerlerin geçerli bir tarih olup olmadığını kontrol eder. Geçerli tarihleri ve tanımlanamayan değerleri farklı renklerle işaretler.", "Inspect date values for invalid or inconsistent entries.")
 Case "HAMU_DateBuilder"
 HAMU_CommandIntro = HAMU_L("Ayrı sütunlarda bulunan yıl, ay ve gün değerlerini birleştirerek gerçek Excel tarihleri oluşturur. Ay alanında sayı veya Türkçe ay adı kullanılabilir.", "Create dates from separate year, month and day fields.")
 Case "HAMU_SheetIndex"
 HAMU_CommandIntro = HAMU_L("Görünür çalışma sayfalarını listeleyen ve her sayfaya tıklanabilir bağlantı veren bir İçindekiler sayfası oluşturur.", "Create a worksheet index with hyperlinks to the workbook's sheets.")
 Case "HAMU_SortSheets"
 HAMU_CommandIntro = HAMU_L("Çalışma kitabındaki çalışma sayfalarını adlarına göre alfabetik olarak yeniden sıralar.", "Sort worksheet tabs alphabetically.")
 Case "HAMU_ShowHidden"
 HAMU_CommandIntro = HAMU_L("Aktif çalışma kitabındaki gizli çalışma sayfalarını, satırları ve sütunları görünür hale getirir.", "Reveal hidden sheets, rows or columns as supported by the command.")
 Case "HAMU_DynamicValidation"
 HAMU_CommandIntro = HAMU_L("Seçili hücrelere liste tipi veri doğrulama uygular. Kaynak olarak adlandırılmış aralık, Tablo[Sütun], doğrudan aralık veya A1# biçimindeki dinamik taşan dizi kullanılabilir. Bir adlandırma A1# kaynağına bağlıysa taşan aralık büyüdükçe doğrulama listesi de dinamik kalır.", "Create list validation using a named range, Table[Column], direct range or spilled range such as A1#. Table/spill sources can expand dynamically.")
 Case "HAMU_SelectionToPNG"
 HAMU_CommandIntro = HAMU_L("Seçili alanı PNG, JPG veya GIF olarak kaydeder. Çok büyük görsel alanlarını durdurur ve geçici grafiği temizler.", "Save the selected range as PNG, JPG or GIF. Reject oversized images and clean up the temporary chart.")
 Case "HAMU_InsertPhotosFromFolder"
 HAMU_CommandIntro = HAMU_L("Seçtiğiniz klasördeki görselleri aktif hücreden başlayarak satırlara ekler ve yan hücreye dosya adını yazar.", "Insert supported images from a folder into cells or over cells. In-cell conversion depends on Excel version support; over-cell images are constrained to cell bounds.")
 Case "HAMU_CreateQRCode"
 HAMU_CommandIntro = HAMU_L("Tek ekranda Karekod veya Barkod seçin; tür değişince model listesi güncellenir. Metin, bağlantı, e-posta ve telefon içeriği; tek giriş veya tek sütundan toplu üretim. Kodlar verilerin sağındaki ilk boş sütuna, kaynak satırlarıyla hizalı eklenir. İnternet gerekir; en fazla 250 kod.", "Select QR Code or Barcode in one form; model choices follow the type. Supports text, links, email, phone, single values or a batch column. Batch codes go in the first empty column to the right, aligned with their source rows. Internet required; at most 250 codes.")
 Case "HAMU_CompareSheets"
 HAMU_CommandIntro = HAMU_L("İki çalışma sayfasını belirlediğiniz anahtar sütuna göre karşılaştırır ve eklenen, silinen veya değişen kayıtları ayrı bir raporda gösterir.", "Compare worksheets and identify differences.")
 Case "HAMU_MarkProblemCells"
 HAMU_CommandIntro = HAMU_L("Sayı gibi görünen metinleri ve eşittir işaretiyle başlayan ancak gerçek formül olmayan metinleri bularak hücreleri farklı renklerle işaretler.", "Highlight cells with the problems selected in the dialog.")
 Case "HAMU_HighlightByCriteria"
 HAMU_CommandIntro = HAMU_L("Seçili alanda sayı, metin, hata veya boş hücreleri seçtiğiniz kritere göre vurgular.", "Highlight cells matching a selected type or criterion.")
 Case "HAMU_HighlightDuplicates"
 HAMU_CommandIntro = HAMU_L("Seçili alanda birden fazla kez bulunan değerleri tespit eder ve tekrar eden hücreleri renklendirir.", "Highlight repeated values within the selected area.")
 Case "HAMU_AddDuplicateCount"
 HAMU_CommandIntro = HAMU_L("Giriş listesindeki her değerin kaç kez geçtiğini hesaplayıp ayrı seçtiğiniz çıkış alanına yazar. Tek çıkış hücresi seçerseniz sonuç aşağı doğru genişler.", "Count occurrences of each input value and write the counts to a separately selected output range. A single output cell defines the starting point.")
 Case "HAMU_ListLinks"
 HAMU_CommandIntro = HAMU_L("Çalışma kitabındaki Excel türü dış bağlantıları yeni bir sayfada listeler.", "List external Excel workbook links on a new worksheet.")
 Case "HAMU_BreakLinks"
 HAMU_CommandIntro = HAMU_L("Çalışma kitabındaki harici Excel bağlantılarını kırar. Bağlantılı formüller mevcut değerlerine dönüştürülebileceği için işlemden önce yedek alınması önerilir.", "Break external Excel links. Linked formulas may become values; back up the workbook before proceeding.")
 Case "HAMU_Help"
 HAMU_CommandIntro = HAMU_L("HAMU Tools içindeki komutların kullanım amacı ve temel kullanım bilgilerini gösterir.", "Open the usage guide with workflows, selection/output explanations, limitations and the complete tool catalog.")
 Case "HAMU_About"
 HAMU_CommandIntro = HAMU_L("HAMU Tools V1 sürümü, HAMU markası ve huseyinavniuzun.com bilgileri ile eklentinin kapsamını gösterir.", "Show HAMU Tools V1, Hüseyin Avni UZUN and the website, GitHub and email links.")
 Case "HAMU_AutoFit"
 HAMU_CommandIntro = HAMU_L("Seçili alanın satır yüksekliğini metne göre sığdırır; sütun genişliği değişmez.", "Fit selected row heights without changing column widths.")
 Case "HAMU_AutoTable"
 HAMU_CommandIntro = HAMU_L("Tek hücre seçildiyse çevresindeki kesintisiz veri alanını, aksi halde seçilen alanı başlıklı Excel tablosuna dönüştürür. HAMU stili ve filtreleri uygular.", "Create an Excel table from a header/data range and apply a consistent table style. A single selected cell uses its current data region.")
 Case "HAMU_ResizeTable"
 HAMU_CommandIntro = HAMU_L("Seçili tablonun altındaki kesintisiz veriyi aynı sütun genişliğinde tabloya dahil eder. Başka tablolarla çakışmayı önler.", "Extend the active Excel table into adjacent data. Select a cell inside an existing table first.")
 Case "HAMU_UDFHelp"
 HAMU_CommandIntro = HAMU_L("HAMU çalışma sayfası fonksiyonlarını ve formül örneklerini gösterir.", "Show the eight HAMU worksheet functions with examples. Function names remain unchanged when switching language.")
 Case "UDF_URLDecode"
 HAMU_CommandIntro = HAMU_L("=HAMU_URLCoz(A1) | UTF-8 URL karakterlerini çözer. Geçersiz yüzde dizisini korur.", "=HAMU_URLCoz(A1) | Decode UTF-8 URL characters; preserve invalid percent sequences.")
 Case "UDF_RenklileriTopla"
 HAMU_CommandIntro = HAMU_L("=HAMU_RengeGoreTopla(A1:A20;B1) | B1 ile aynı dolgu rengine sahip sayıları toplar. Biçim değişimi sonrası F9 kullanın.", "=HAMU_RengeGoreTopla(A1:A20;B1) | Sum numbers with the same direct fill color as B1. Recalculate with F9 after changing formatting.")
 Case "UDF_RengeGoreSay"
 HAMU_CommandIntro = HAMU_L("=HAMU_RengeGoreSay(A1:A20;B1) | B1 ile aynı dolgu rengine sahip hücreleri sayar. Biçim değişimi sonrası F9 kullanın.", "=HAMU_RengeGoreSay(A1:A20;B1) | Count cells with the same direct fill color as B1. Recalculate with F9 after changing formatting.")
 Case "UDF_YerelDosyaYolu"
 HAMU_CommandIntro = HAMU_L("=HAMU_YerelDosyaYolu(A1) | Kişisel OneDrive URL yolunu yerel OneDrive ortam değişkeniyle eşler; diğer URL türlerini korur.", "=HAMU_YerelDosyaYolu(A1) | Map a personal OneDrive URL to a local path using the local OneDrive environment variable; preserve other URL types.")
 Case "UDF_GizliLink"
 HAMU_CommandIntro = HAMU_L("=HAMU_GizliLink(A1) | Hücrenin ilk köprü adresini getirir.", "=HAMU_GizliLink(A1) | Return the first hyperlink address in the cell.")
 Case "UDF_LinkParamDegeri"
 HAMU_CommandIntro = HAMU_L("=HAMU_LinkParamDegeri(A1;""id"") | URL içindeki sorgu parametresinin değerini getirir.", "=HAMU_LinkParamDegeri(A1;""id"") | Return the value of the specified URL query parameter.")
 Case "UDF_SayidanMetine"
 HAMU_CommandIntro = HAMU_L("=HAMU_SayidanMetine(123) | Türkçe: yüz yirmi üç. 15 basamaktan küçük mutlak değerler; ondalık kısmı tek tek okur.", "=HAMU_SayidanMetine(123) | Spell numbers in Turkish. Supports absolute values below 15 digits and reads decimal digits individually. The output language stays Turkish.")
 Case "UDF_TablodanVeriGetir"
 HAMU_CommandIntro = HAMU_L("=HAMU_TablodanVeriGetir(A1;""Tablo1"";""Kod"";""Ad"") | Formülün bulunduğu kitapta tablodan ilk eşleşen değeri getirir.", "=HAMU_TablodanVeriGetir(A1;""Tablo1"";""Kod"";""Ad"") | Return the first matching value from a table in the workbook containing the formula.")
 Case "HAMU_CountByColor"
 HAMU_CommandIntro = HAMU_L("Seçilen aralıkta örnek hücrenin görüntülenen dolgu rengiyle eşleşen hücreleri sayar; koşullu biçimlendirme dahildir.", "Count cells matching the sample cell's displayed fill color, including conditional formatting.")
 Case "HAMU_SumByColor"
 HAMU_CommandIntro = HAMU_L("Örnek hücrenin görüntülenen dolgu rengiyle eşleşen sayısal değerleri toplar; metin, boş ve hata değerlerini toplama katmaz.", "Sum numeric values matching the sample cell's displayed fill color, including conditional formatting. Ignore blank, text and error values.")
 Case "HAMU_FillBlanksZero"
 HAMU_CommandIntro = HAMU_L("Seçili aralıktaki gerçekten boş hücrelere 0 yazar. Dolu hücreleri ve boş metin döndüren formülleri korur.", "Write zero only to genuinely empty cells. Preserve populated cells and formulas returning empty text.")
 Case "HAMU_AddTextAffixes"
 HAMU_CommandIntro = HAMU_L("Seçili tek sütundaki metinlerin başına ve/veya sonuna metin ekler. Sonuç seçtiğiniz başlangıç hücresine yazılır; kaynak korunur. Boş hücreler boş kalır.", "Add text before or after values and write to a separate output range. Preserve the source and leading zeros.")
 Case "HAMU_LeadingZeros"
 HAMU_CommandIntro = HAMU_L("Kodları sabit uzunluğa baştan sıfır ekleyerek düzenler veya baştaki sıfırları kaldırır. Sonuç metin olarak seçtiğiniz hedefe yazılır; uzun kodlar ve kaynak korunur.", "Add or remove leading zeros from text codes. Preserve source values and write the result as text.")
 Case "HAMU_DeduplicateItems"
 HAMU_CommandIntro = HAMU_L(("Virgül veya belirttiğiniz ayraçla ayrılmış hücre içi öğeleri tekilleştirir. Örneğin elma, armut, elma " & ChrW(8594) & " elma, armut. İlk sıra korunur; sonuç ayrı hedefe yazılır."), "Remove repeated items within each cell using the chosen separator, keeping their first occurrence.")
 Case "HAMU_ExtractTextPart"
 HAMU_CommandIntro = HAMU_L("Soldan/sağdan belirli sayıda karakteri, bir işaretin öncesini/sonrasını veya iki işaret arasını ayıklar. İlk eşleşmeyi kullanır; eşleşme yoksa boş sonuç verir. Kaynak korunur.", "Extract characters from the left/right or text before, after or between specified markers.")
 Case "HAMU_DateRangeList"
 HAMU_CommandIntro = HAMU_L("Başlangıç ve bitiş dahil tarih listesi oluşturur. Tüm günler veya yalnız hafta içi seçilir; hafta içi seçeneği resmî tatilleri hesaplamaz. Sonuç seçtiğiniz başlangıç hücresine yazılır.", "Generate dates between a start and end date, optionally excluding weekends. Choose the output start cell.")
 Case "HAMU_FormulaAudit"
 HAMU_CommandIntro = HAMU_L("Kitaptaki formülleri kaynak hücre bağlantısı, formül metni, görüntülenen sonuç, hata durumu ve dış kitap başvurusu bilgisiyle yeni sayfada listeler. Kaynak formüller değiştirilmez. Dış başvuru tespiti doğrudan XLS/CSV kitap başvurularını tarar.", "List formulas, source-cell links, displayed results, errors and direct external workbook references on a new sheet. Source formulas remain unchanged.")
 Case "HAMU_SelectSpecialCells"
 HAMU_CommandIntro = HAMU_L("Mevcut seçimdeki formül, sabit değer, hata, boş veya görünür hücreleri seçer. Değer ve biçimler değiştirilmez; boş metin döndüren formüller gerçek boş sayılmaz.", "Select formulas, constants, errors, blanks or visible cells within the current area. Do not change values or formatting.")
 Case "HAMU_ListDefinedNames"
 HAMU_CommandIntro = HAMU_L("Kitap ve sayfa kapsamındaki tanımlı adları, başvurularını, görünürlüklerini ve geçersiz başvuru durumlarını yeni sayfada listeler. Adları değiştirmez veya silmez.", "List workbook/worksheet names, references, visibility and broken references on a new sheet without modifying the names.")
 Case "HAMU_DataEntry"
 HAMU_CommandIntro = HAMU_L("Excel’in yerleşik veri giriş formunu açar. Bir tabloya veya benzersiz başlıkları olan kesintisiz listeye tıklayın; en fazla 32 sütun.", "Open Excel's native data form for the current table or contiguous list with unique headers; at most 32 columns.")
 Case "HAMU_ClearHighlights"
 HAMU_CommandIntro = HAMU_L("Seçili alandaki HAMU vurgu renklerini kaldırır. Aynı oturumda önceki dolguyu geri yükler; daha eski vurgular sarı/turuncu paletiyle tanınır.", "Remove HAMU highlights in the selection. Restore prior fills in the same session; older highlights are recognized by their yellow/orange palette.")
 Case "UDF_LambdaMultiLookup"
 HAMU_CommandIntro = HAMU_L("İki ayrı anahtar/sonuç sütununda tam eşleşme arar; ilk sonucu verir. Örnek: =HAMU_L_CokluAra(G2;A2:A10;B2:B10;D2:D10;E2:E10). LAMBDA şablonundan aktarılır; Microsoft 365/Excel 2024 gerekir.", "Exact lookup across two key/result pairs, returning the first match. Example: =HAMU_L_CokluAra(G2;A2:A10;B2:B10;D2:D10;E2:E10). Transfer it from the LAMBDA template; Microsoft 365/Excel 2024 required.")
 Case "UDF_LambdaMultiSum"
 HAMU_CommandIntro = HAMU_L("İki ayrı anahtar/değer sütununda koşula uyan sayıları toplar. Örnek: =HAMU_L_CokluTopla(G2;A2:A10;B2:B10;D2:D10;E2:E10). Şablondan aktarılır; Microsoft 365/Excel 2024 gerekir.", "Sum matching values across two key/value pairs. Example: =HAMU_L_CokluTopla(G2;A2:A10;B2:B10;D2:D10;E2:E10). Transfer from the template; Microsoft 365/Excel 2024 required.")
 Case "UDF_LambdaUniqueMerge"
 HAMU_CommandIntro = HAMU_L("İki alandaki dolu değerleri tek bir sıralı benzersiz listeye dönüştürür. Örnek: =HAMU_L_BenzersizBirlestir(A2:B10;D2:E10). Sonuç için boş taşma alanı gerekir. Microsoft 365/Excel 2024 gerekir.", "Combine nonblank values from two areas into a sorted unique list. Example: =HAMU_L_BenzersizBirlestir(A2:B10;D2:E10). Requires a blank spill area and Microsoft 365/Excel 2024.")
 Case "HAMU_SaveCopy"
 HAMU_CommandIntro = HAMU_L("Dosyanın bir kopyasını seçtiğiniz konuma kaydeder. Açık dosyanın adı, konumu ve kaydedilmemiş değişiklik durumu korunur; biçim ve makrolar değiştirilmez.", "Save a copy at a chosen location, preserving the open workbook identity, unsaved state, file format and macros.")
 Case "HAMU_ExpandColumns"
 HAMU_CommandIntro = HAMU_L("Seçili alanın sütun genişliğini metne göre ayarlar.", "Fit selected column widths to their content.")
 Case "HAMU_TableTotals"
 HAMU_CommandIntro = HAMU_L("Etkin tablonun toplam satırını açar veya kapatır. Sayısal sütunlarda toplam gösterir.", "Toggle the active table total row and sum numeric columns.")
 Case "HAMU_CopyVisibleTable"
 HAMU_CommandIntro = HAMU_L("Filtre sonrası görünen kayıtları yeni sayfada bağımsız bir tabloya kopyalar. Kaynak değişmez.", "Copy visible filtered records into a standalone table on a new sheet.")
 End Select
End Function
Public Function HAMU_CommandLabel(ByVal id As String) As String
 Select Case id
 Case "HAMU_Transfer"
 HAMU_CommandLabel = HAMU_L("Aktarma Sihirbazı", "Transfer Wizard")
 Case "HAMU_Backup"
 HAMU_CommandLabel = HAMU_L("Yedek Oluştur", "Create Backup")
 Case "HAMU_AddDateToFilename"
 HAMU_CommandLabel = HAMU_L("Tarihli Sürüm", "Dated Version")
 Case "HAMU_ExportSheetsAsPDF"
 HAMU_CommandLabel = HAMU_L("PDF Kaydet", "Save PDF")
 Case "HAMU_FolderIndex"
 HAMU_CommandLabel = HAMU_L("Klasör İçeriği", "Folder Contents")
 Case "HAMU_CombineCSVs"
 HAMU_CommandLabel = HAMU_L("CSV Dosyalarını Birleştir", "Combine CSV Files")
 Case "HAMU_ConsolidateSheets"
 HAMU_CommandLabel = HAMU_L("Sayfaları Birleştir", "Combine Sheets")
 Case "HAMU_AppendSelectedTables"
 HAMU_CommandLabel = HAMU_L("Tabloları Birleştir", "Combine Tables")
 Case "HAMU_SmartAppendByHeaders"
 HAMU_CommandLabel = HAMU_L("Akıllı Birleştir", "Smart Append")
 Case "HAMU_CollectFilesData"
 HAMU_CommandLabel = HAMU_L("Dosyalardan Topla", "Collect Files")
 Case "HAMU_SmartJoin"
 HAMU_CommandLabel = HAMU_L("Anahtara Göre Tablo Eşleştir", "Join by Key")
 Case "HAMU_GroupAndSummarize"
 HAMU_CommandLabel = HAMU_L("Grupla ve Özetle", "Group and Summarize")
 Case "HAMU_UnpivotColumnsToRows"
 HAMU_CommandLabel = HAMU_L("Unpivot", "Unpivot")
 Case "HAMU_SplitDataWizard"
 HAMU_CommandLabel = HAMU_L("Veriyi Böl", "Split Data")
 Case "HAMU_SplitByValueToFiles"
 HAMU_CommandLabel = HAMU_L("Dosyalara Böl", "Split to Files")
 Case "HAMU_SplitURLParameters"
 HAMU_CommandLabel = HAMU_L("URL Ayıkla", "Extract URL")
 Case "HAMU_PivotExport"
 HAMU_CommandLabel = HAMU_L("Pivotu Böl", "Split Pivot")
 Case "HAMU_PivotRowsToColumns"
 HAMU_CommandLabel = HAMU_L("Pivot", "Pivot")
 Case "HAMU_FillDownBlanks"
 HAMU_CommandLabel = HAMU_L("Boşlukları Aşağı Doldur", "Fill Down Blanks")
 Case "HAMU_MergeDuplicates"
 HAMU_CommandLabel = HAMU_L("Tekrarları Birleştir", "Merge Duplicate Records")
 Case "HAMU_StandardizeHeaders"
 HAMU_CommandLabel = HAMU_L("Başlıkları Düzelt", "Clean Headers")
 Case "HAMU_ConvertColumnTypes"
 HAMU_CommandLabel = HAMU_L("Veri Türünü Düzelt", "Convert Types")
 Case "HAMU_SelectColumnsToNewTable"
 HAMU_CommandLabel = HAMU_L("Sütun Kopyala", "Copy Columns")
 Case "HAMU_FilterRowsToNewTable"
 HAMU_CommandLabel = HAMU_L("Koşulla Satır Filtrele", "Filter Rows")
 Case "HAMU_SampleData"
 HAMU_CommandLabel = HAMU_L("Veriden Örnek Al", "Sample Data")
 Case "HAMU_ShuffleRows"
 HAMU_CommandLabel = HAMU_L("Satırları Karıştır", "Shuffle Rows")
 Case "HAMU_CompareTwoLists"
 HAMU_CommandLabel = HAMU_L("İki Listeyi Karşılaştır", "Compare Lists")
 Case "HAMU_ExtractUniqueList"
 HAMU_CommandLabel = HAMU_L("Benzersiz Liste", "Unique List")
 Case "HAMU_DataProfile"
 HAMU_CommandLabel = HAMU_L("Veri Profili", "Data Profile")
 Case "HAMU_DataQualityReport"
 HAMU_CommandLabel = HAMU_L("Kalite Raporu", "Quality Report")
 Case "HAMU_FindMissingCombinations"
 HAMU_CommandLabel = HAMU_L("Eksikleri Bul", "Find Missing")
 Case "HAMU_CleanData"
 HAMU_CommandLabel = HAMU_L("Temizle Kırp", "Clean Trim")
 Case "HAMU_TextCase"
 HAMU_CommandLabel = HAMU_L("Harf Dönüştür", "Change Case")
 Case "HAMU_PatternFind"
 HAMU_CommandLabel = HAMU_L("Normal İfade ile Bul", "Find by Regular Expression")
 Case "HAMU_RegexReplace"
 HAMU_CommandLabel = HAMU_L("Normal İfade ile Değiştir", "Replace by Regular Expression")
 Case "HAMU_ExtractDigits"
 HAMU_CommandLabel = HAMU_L("Metinden Sayıları Ayıkla", "Extract Digits")
 Case "HAMU_ConvertQuotedNumbers"
 HAMU_CommandLabel = HAMU_L("Metin Olarak Saklanan Sayıları Düzelt", "Convert Text Numbers")
 Case "HAMU_MaskData"
 HAMU_CommandLabel = HAMU_L("Veriyi Maskele", "Mask Data")
 Case "HAMU_ConvertFormulasToValues"
 HAMU_CommandLabel = HAMU_L("Formülleri Değere Dönüştür", "Formulas to Values")
 Case "HAMU_UnmergeAndFill"
 HAMU_CommandLabel = HAMU_L("Birleşimleri Çöz ve Doldur", "Unmerge and Fill")
 Case "HAMU_StandardFormat"
 HAMU_CommandLabel = HAMU_L("Standart Sayı ve Tarih Biçimi", "Number and Date Format")
 Case "HAMU_DeleteBlankCells"
 HAMU_CommandLabel = HAMU_L("Boşlukları Sil", "Delete Blanks")
 Case "HAMU_DeleteDropdowns"
 HAMU_CommandLabel = HAMU_L("Eski Dropdown Nesnelerini Sil", "Remove Old Dropdown Objects")
 Case "HAMU_RowStriping"
 HAMU_CommandLabel = HAMU_L("Satırları Renklendir", "Stripe Rows")
 Case "HAMU_DatePicker"
 HAMU_CommandLabel = HAMU_L("Tarih Seçici", "Date Picker")
 Case "HAMU_FixDates"
 HAMU_CommandLabel = HAMU_L("Tarihleri Düzelt", "Fix Dates")
 Case "HAMU_CheckDates"
 HAMU_CommandLabel = HAMU_L("Tarih Kontrolü", "Check Dates")
 Case "HAMU_DateBuilder"
 HAMU_CommandLabel = HAMU_L("Tarih Oluştur", "Build Dates")
 Case "HAMU_SheetIndex"
 HAMU_CommandLabel = HAMU_L("Sayfa İndeksi", "Sheet Index")
 Case "HAMU_SortSheets"
 HAMU_CommandLabel = HAMU_L("Sayfaları Alfabetik Sırala", "Sort Sheets")
 Case "HAMU_ShowHidden"
 HAMU_CommandLabel = HAMU_L("Tüm Gizlileri Göster", "Show Hidden Items")
 Case "HAMU_DynamicValidation"
 HAMU_CommandLabel = HAMU_L("Dinamik Veri Doğrulama Listesi", "Dynamic Validation List")
 Case "HAMU_SelectionToPNG"
 HAMU_CommandLabel = HAMU_L("Resim Kaydet", "Save Image")
 Case "HAMU_InsertPhotosFromFolder"
 HAMU_CommandLabel = HAMU_L("Toplu Fotoğraf", "Batch Photos")
 Case "HAMU_CreateQRCode"
 HAMU_CommandLabel = HAMU_L("Karekod ve Barkod", "QR and Barcode")
 Case "HAMU_CompareSheets"
 HAMU_CommandLabel = HAMU_L("Sayfaları Karşılaştır", "Compare Sheets")
 Case "HAMU_MarkProblemCells"
 HAMU_CommandLabel = HAMU_L("Sorunlu Hücreler", "Problem Cells")
 Case "HAMU_HighlightByCriteria"
 HAMU_CommandLabel = HAMU_L("Hücre Türüne Göre Vurgula", "Highlight Cell Types")
 Case "HAMU_HighlightDuplicates"
 HAMU_CommandLabel = HAMU_L("Tekrar Edenleri Vurgula", "Highlight Duplicates")
 Case "HAMU_AddDuplicateCount"
 HAMU_CommandLabel = HAMU_L("Tekrar Sayısını Yaz", "Write Duplicate Counts")
 Case "HAMU_ListLinks"
 HAMU_CommandLabel = HAMU_L("Harici Bağlantıları Listele", "List External Links")
 Case "HAMU_BreakLinks"
 HAMU_CommandLabel = HAMU_L("Harici Bağlantıları Kır", "Break External Links")
 Case "HAMU_Help"
 HAMU_CommandLabel = HAMU_L("Yardım", "Help")
 Case "HAMU_About"
 HAMU_CommandLabel = HAMU_L("Hakkında", "About")
 Case "HAMU_AutoFit"
 HAMU_CommandLabel = HAMU_L("Sığdır", "AutoFit")
 Case "HAMU_AutoTable"
 HAMU_CommandLabel = HAMU_L("Otomatik Tablo Oluştur", "Create Table")
 Case "HAMU_ResizeTable"
 HAMU_CommandLabel = HAMU_L("Tabloyu Veriye Genişlet", "Expand Table")
 Case "HAMU_UDFHelp"
 HAMU_CommandLabel = HAMU_L("Çalışma Sayfası Fonksiyonları", "Worksheet Functions")
 Case "UDF_URLDecode"
 HAMU_CommandLabel = HAMU_L("HAMU_URLCoz", "HAMU_URLCoz")
 Case "UDF_RenklileriTopla"
 HAMU_CommandLabel = HAMU_L("HAMU_RengeGoreTopla", "HAMU_RengeGoreTopla")
 Case "UDF_RengeGoreSay"
 HAMU_CommandLabel = HAMU_L("HAMU_RengeGoreSay", "HAMU_RengeGoreSay")
 Case "UDF_YerelDosyaYolu"
 HAMU_CommandLabel = HAMU_L("HAMU_YerelDosyaYolu", "HAMU_YerelDosyaYolu")
 Case "UDF_GizliLink"
 HAMU_CommandLabel = HAMU_L("HAMU_GizliLink", "HAMU_GizliLink")
 Case "UDF_LinkParamDegeri"
 HAMU_CommandLabel = HAMU_L("HAMU_LinkParamDegeri", "HAMU_LinkParamDegeri")
 Case "UDF_SayidanMetine"
 HAMU_CommandLabel = HAMU_L("HAMU_SayidanMetine", "HAMU_SayidanMetine")
 Case "UDF_TablodanVeriGetir"
 HAMU_CommandLabel = HAMU_L("HAMU_TablodanVeriGetir", "HAMU_TablodanVeriGetir")
 Case "HAMU_CountByColor"
 HAMU_CommandLabel = HAMU_L("Renge Göre Say", "Count by Color")
 Case "HAMU_SumByColor"
 HAMU_CommandLabel = HAMU_L("Renge Göre Topla", "Sum by Color")
 Case "HAMU_FillBlanksZero"
 HAMU_CommandLabel = HAMU_L("Boşlukları Sıfırla Doldur", "Fill Blanks with Zero")
 Case "HAMU_AddTextAffixes"
 HAMU_CommandLabel = HAMU_L("Başa / Sona Ekle", "Add Prefix or Suffix")
 Case "HAMU_LeadingZeros"
 HAMU_CommandLabel = HAMU_L("Baştaki Sıfırlar", "Leading Zeros")
 Case "HAMU_DeduplicateItems"
 HAMU_CommandLabel = HAMU_L("Hücre İçi Tekrarlar", "Deduplicate Cell Items")
 Case "HAMU_ExtractTextPart"
 HAMU_CommandLabel = HAMU_L("Metin Parçası Al", "Extract Text Part")
 Case "HAMU_DateRangeList"
 HAMU_CommandLabel = HAMU_L("Tarih Listesi", "Date List")
 Case "HAMU_FormulaAudit"
 HAMU_CommandLabel = HAMU_L("Formül Denetimi", "Formula Audit")
 Case "HAMU_SelectSpecialCells"
 HAMU_CommandLabel = HAMU_L("Özel Hücreleri Seç", "Select Special Cells")
 Case "HAMU_ListDefinedNames"
 HAMU_CommandLabel = HAMU_L("Tanımlı Adları Listele", "List Defined Names")
 Case "HAMU_DataEntry"
 HAMU_CommandLabel = HAMU_L("Veri Girişi", "Data Entry")
 Case "HAMU_ClearHighlights"
 HAMU_CommandLabel = HAMU_L("Vurguyu Kaldır", "Clear Highlights")
 Case "UDF_LambdaMultiLookup"
 HAMU_CommandLabel = HAMU_L("HAMU_L_CokluAra", "HAMU_L_CokluAra")
 Case "UDF_LambdaMultiSum"
 HAMU_CommandLabel = HAMU_L("HAMU_L_CokluTopla", "HAMU_L_CokluTopla")
 Case "UDF_LambdaUniqueMerge"
 HAMU_CommandLabel = HAMU_L("HAMU_L_BenzersizBirlestir", "HAMU_L_BenzersizBirlestir")
 Case "HAMU_SaveCopy"
 HAMU_CommandLabel = HAMU_L("Kopya Kaydet", "Save Copy")
 Case "HAMU_ExpandColumns"
 HAMU_CommandLabel = HAMU_L("Genişlet", "Fit Width")
 Case "HAMU_TableTotals"
 HAMU_CommandLabel = HAMU_L("Toplam Satırı", "Total Row")
 Case "HAMU_CopyVisibleTable"
 HAMU_CommandLabel = HAMU_L("Görünür Satırları Kopyala", "Copy Visible Rows")
 End Select
End Function
Public Function HAMU_NeedsWorksheet(ByVal id As String) As Boolean
    If Left$(id, 4) = "UDF_" Then Exit Function
    Select Case id
        Case "HAMU_SaveCopy": HAMU_NeedsWorksheet = False
        Case "HAMU_Transfer": HAMU_NeedsWorksheet = False
        Case "HAMU_FormulaAudit", "HAMU_ListDefinedNames": HAMU_NeedsWorksheet = False
        Case "HAMU_Backup": HAMU_NeedsWorksheet = False
        Case "HAMU_AddDateToFilename": HAMU_NeedsWorksheet = False
        Case "HAMU_FolderIndex": HAMU_NeedsWorksheet = False
        Case "HAMU_SheetIndex": HAMU_NeedsWorksheet = False
        Case "HAMU_SortSheets": HAMU_NeedsWorksheet = False
        Case "HAMU_ShowHidden": HAMU_NeedsWorksheet = False
        Case "HAMU_Help": HAMU_NeedsWorksheet = False
        Case "HAMU_About": HAMU_NeedsWorksheet = False
        Case "HAMU_UDFHelp": HAMU_NeedsWorksheet = False
        Case "UDF_URLDecode": HAMU_NeedsWorksheet = False
        Case "UDF_RenklileriTopla": HAMU_NeedsWorksheet = False
        Case "UDF_RengeGoreSay": HAMU_NeedsWorksheet = False
        Case "UDF_YerelDosyaYolu": HAMU_NeedsWorksheet = False
        Case "UDF_GizliLink": HAMU_NeedsWorksheet = False
        Case "UDF_LinkParamDegeri": HAMU_NeedsWorksheet = False
        Case "UDF_SayidanMetine": HAMU_NeedsWorksheet = False
        Case "UDF_TablodanVeriGetir": HAMU_NeedsWorksheet = False
        Case Else: HAMU_NeedsWorksheet = True
    End Select
End Function

