Attribute VB_Name = "modHAMU_UIText"
Option Explicit
Option Compare Text
Option Private Module
Public Function HAMU_UIText(ByVal id As String, ByVal field As String) As String
 Dim key As String
 key = id & "|" & field
 HAMU_UIText = HAMU_UITextPart1(key)
 If Len(HAMU_UIText) > 0 Then Exit Function
 HAMU_UIText = HAMU_UITextPart2(key)
 If Len(HAMU_UIText) > 0 Then Exit Function
 HAMU_UIText = HAMU_UITextPart3(key)
 If Len(HAMU_UIText) > 0 Then Exit Function
 HAMU_UIText = HAMU_UITextPart4(key)
 If Len(HAMU_UIText) > 0 Then Exit Function
 HAMU_UIText = HAMU_UITextPart5(key)
 If Len(HAMU_UIText) > 0 Then Exit Function
 HAMU_UIText = HAMU_UITextPart6(key)
 If Len(HAMU_UIText) > 0 Then Exit Function
 HAMU_UIText = HAMU_UITextPart7(key)
 If Len(HAMU_UIText) > 0 Then Exit Function
 HAMU_UIText = HAMU_UITextPart8(key)
 If Len(HAMU_UIText) > 0 Then Exit Function
 HAMU_UIText = HAMU_UITextPart9(key)
 If Len(HAMU_UIText) > 0 Then Exit Function
End Function

Private Function HAMU_UITextPart1(ByVal key As String) As String
 Select Case key
 Case "tabHAMU|label"
 HAMU_UITextPart1 = HAMU_L("HAMU Tools", "HAMU Tools")
 Case "g0|label"
 HAMU_UITextPart1 = HAMU_L("Aktarma", "Transfer")
 Case "r_HAMU_Transfer|label"
 HAMU_UITextPart1 = HAMU_L("Aktar", "Transfer")
 Case "r_HAMU_Transfer|screentip"
 HAMU_UITextPart1 = HAMU_L("Aktarma Sihirbazı", "Transfer Wizard")
 Case "r_HAMU_Transfer|supertip"
 HAMU_UITextPart1 = HAMU_L("Şablondan LAMBDA, hücre stili, adlandırılmış veri aralığı, sayfa verisi ve Excel tablosu aktarır. Sayfa/tablo aktarımı değer ve biçim kopyalar; VBA taşımaz.", "Transfer LAMBDA definitions, cell styles, named data ranges, sheet data and Excel tables. Sheet/table transfer copies values and formatting, without VBA.")
 Case "g1|label"
 HAMU_UITextPart1 = HAMU_L("Dosya", "Files")
 Case "r_HAMU_AddDateToFilename|label"
 HAMU_UITextPart1 = HAMU_L("Tarihli Sürüm", "Dated Version")
 Case "r_HAMU_AddDateToFilename|screentip"
 HAMU_UITextPart1 = HAMU_L("Tarihli Sürüm", "Dated Version")
 Case "r_HAMU_AddDateToFilename|supertip"
 HAMU_UITextPart1 = HAMU_L("Kaydedilmiş dosyanın aynı klasörde tarih ve saat eklenmiş kopyasını oluşturur. Yeni kopya açılıp doğrulandıktan sonra eski dosya kapatılır ve silinir. Yeni dosya açılamazsa eski dosya korunur.", "Create a timestamped copy in the original folder. Open and verify the copy before closing and deleting the original; keep the original if verification fails.")
 Case "r_HAMU_FolderIndex|label"
 HAMU_UITextPart1 = HAMU_L("Klasör İçeriği", "Folder Contents")
 Case "r_HAMU_FolderIndex|screentip"
 HAMU_UITextPart1 = HAMU_L("Klasör İçeriği", "Folder Contents")
 Case "r_HAMU_FolderIndex|supertip"
 HAMU_UITextPart1 = HAMU_L("Klasör seçildikten sonra yalnız bu klasör veya alt klasörlerle birlikte tarama seçilir. Alt klasör taraması uzun sürebilir. Dosyaları bağlantı, yol, tür, okunabilir boyut ve tarihlerle listeler.", "List files with hyperlinks, local paths, extensions, types, readable sizes and creation/modification dates. Choose current-folder or recursive scanning; large folder trees take longer.")
 Case "g2|label"
 HAMU_UITextPart1 = HAMU_L("Veri", "Data")
 Case "r_HAMU_UnpivotColumnsToRows|label"
 HAMU_UITextPart1 = HAMU_L("Unpivot", "Unpivot")
 Case "r_HAMU_UnpivotColumnsToRows|screentip"
 HAMU_UITextPart1 = HAMU_L("Sütunları Satırlara Dönüştür", "Unpivot")
 Case "r_HAMU_UnpivotColumnsToRows|supertip"
 HAMU_UITextPart1 = HAMU_L("Solda belirlediğiniz birden fazla sabit sütunu korur; sağa doğru uzanan pivot sütunlarını Alan ve Değer yapısında satırlara dönüştürür.", "Convert wide columns into field/value rows. Keep text headers such as 2026.01 as text.")
 Case "r_HAMU_DataProfile|label"
 HAMU_UITextPart1 = HAMU_L("Veri Profili", "Data Profile")
 Case "r_HAMU_DataProfile|screentip"
 HAMU_UITextPart1 = HAMU_L("Veri Profilini Çıkar", "Data Profile")
 Case "r_HAMU_DataProfile|supertip"
 HAMU_UITextPart1 = HAMU_L("Seçtiğiniz başlıklı veri alanındaki sütunların doluluk, boşluk, benzersiz değer ve veri türlerini özetler. Sorunları görmek için bir rapor oluşturur; kaynak veriyi değiştirmez.", "Summarize the contents and data types of selected columns to help assess a dataset.")
 Case "r_HAMU_DataQualityReport|label"
 HAMU_UITextPart1 = HAMU_L("Kalite Raporu", "Quality Report")
 Case "r_HAMU_DataQualityReport|screentip"
 HAMU_UITextPart1 = HAMU_L("Veri Kalitesi Raporu", "Data Quality")
 Case "r_HAMU_DataQualityReport|supertip"
 HAMU_UITextPart1 = HAMU_L("Başlıklı tabloyu denetleyerek boş alan, tekrar eden kayıt ve veri türü sorunlarını raporlar. Sonuç yeni sayfaya yazılır; sorunları kaynaktan otomatik silmez. Önce raporu inceleyin.", "Inspect missing values, duplicates and data consistency. Review the report before deciding how to clean the source data.")
 Case "r_HAMU_FindMissingCombinations|label"
 HAMU_UITextPart1 = HAMU_L("Eksikleri Bul", "Find Missing")
 Case "r_HAMU_FindMissingCombinations|screentip"
 HAMU_UITextPart1 = HAMU_L("Eksik Kombinasyonları Bul", "Missing Combinations")
 Case "r_HAMU_FindMissingCombinations|supertip"
 HAMU_UITextPart1 = HAMU_L("Mağaza, Ürün, Hafta gibi seçtiğiniz sütunların benzersiz değerlerinden tüm olası kombinasyonları oluşturur ve kaynak veri setinde bulunmayan kombinasyonları yeni bir sayfada listeler.", "Find missing combinations of selected category columns.")
 Case "r_HAMU_PivotRowsToColumns|label"
 HAMU_UITextPart1 = HAMU_L("Pivot", "Pivot")
 Case "r_HAMU_PivotRowsToColumns|screentip"
 HAMU_UITextPart1 = HAMU_L("Satırlardan Çapraz Tablo Oluştur", "Pivot")
 Case "r_HAMU_PivotRowsToColumns|supertip"
 HAMU_UITextPart1 = HAMU_L("Uzun formattaki veriyi tekrar geniş çapraz tabloya dönüştürür. Birden fazla sabit anahtar sütunu kullanabilir; yeni sütun başlıklarını seçilen alandan üretir ve aynı hücreye düşen kayıtları toplama, ilk değer veya sayım yöntemiyle birleştirir.", "Build a cross-tab by moving row categories into columns and summarizing values.")
 Case "r_HAMU_StandardizeHeaders|label"
 HAMU_UITextPart1 = HAMU_L("Başlıkları Düzelt", "Clean Headers")
 Case "r_HAMU_StandardizeHeaders|screentip"
 HAMU_UITextPart1 = HAMU_L("Başlıkları Standardize Et", "Normalize Headers")
 Case "r_HAMU_StandardizeHeaders|supertip"
 HAMU_UITextPart1 = HAMU_L("Sütun başlıklarını teknik snake_case biçimine dönüştürebilir veya gereksiz boşlukları temizleyerek okunabilir başlık yapısını koruyabilir. Aynı başlık oluşursa otomatik olarak benzersiz ad üretir.", "Clean spaces and normalize column headings to a consistent naming scheme.")
 Case "r_HAMU_CombineCSVs|label"
 HAMU_UITextPart1 = HAMU_L("CSV Dosyalarını Birleştir", "Combine CSV Files")
 Case "r_HAMU_CombineCSVs|screentip"
 HAMU_UITextPart1 = HAMU_L("CSV Dosyalarını Birleştir", "Combine CSV Files")
 Case "r_HAMU_CombineCSVs|supertip"
 HAMU_UITextPart1 = HAMU_L("Seçtiğiniz klasördeki CSV dosyalarını tek sayfada alt alta toplar; ilk dosyanın başlığını kullanır ve sonraki dosyalarda başlık satırını atlar.", "Append CSV files from a folder to one sheet. Keep the first file's header and skip header rows in subsequent files.")
 Case "r_HAMU_ConsolidateSheets|label"
 HAMU_UITextPart1 = HAMU_L("Sayfaları Birleştir", "Combine Sheets")
 Case "r_HAMU_ConsolidateSheets|screentip"
 HAMU_UITextPart1 = HAMU_L("Sayfaları Alt Alta Birleştir", "Append Worksheets")
 Case "r_HAMU_ConsolidateSheets|supertip"
 HAMU_UITextPart1 = HAMU_L("Aktif çalışma kitabındaki sayfaların kullanılan alanlarını tek sayfada birleştirir ve kaynak sayfa bilgisini sonuç verisine ekler.", "Append used ranges from worksheets in the active workbook and add the source worksheet name to the results.")
 Case "r_HAMU_AppendSelectedTables|label"
 HAMU_UITextPart1 = HAMU_L("Tabloları Birleştir", "Combine Tables")
 Case "r_HAMU_AppendSelectedTables|screentip"
 HAMU_UITextPart1 = HAMU_L("Seçili Tabloları Alt Alta Birleştir", "Append Selected Tables")
 Case "r_HAMU_AppendSelectedTables|supertip"
 HAMU_UITextPart1 = HAMU_L("Belirttiğiniz Excel tablolarının aynı yapıdaki satırlarını tek bir sonuç tablosunda alt alta toplar.", "Combine rows from selected Excel tables with matching structures into one result table.")
 Case "r_HAMU_CollectFilesData|label"
 HAMU_UITextPart1 = HAMU_L("Dosyalardan Topla", "Collect Files")
 Case "r_HAMU_CollectFilesData|screentip"
 HAMU_UITextPart1 = HAMU_L("Klasörden Veri Topla", "Collect Folder Data")
 Case "r_HAMU_CollectFilesData|supertip"
 HAMU_UITextPart1 = HAMU_L("Klasördeki Excel ve CSV dosyalarının ilk veri sayfalarını tek sonuçta alt alta toplar. İlk satır başlıktır; aynı başlıklar eşleştirilir. KaynakDosya eklenir, kaynak dosyalar değiştirilmez. Aynı tür tabloları içeren dosyalar kullanın.", "Read the first data sheet from Excel/CSV files in a folder and append matching headers. Add the source filename without changing source files.")
 Case "r_HAMU_SmartJoin|label"
 HAMU_UITextPart1 = HAMU_L("Anahtara Göre Tablo Eşleştir", "Join by Key")
 Case "r_HAMU_SmartJoin|screentip"
 HAMU_UITextPart1 = HAMU_L("Anahtara Göre Tablo Eşleştir", "Join by Key")
 End Select
End Function

Private Function HAMU_UITextPart2(ByVal key As String) As String
 Select Case key
 Case "r_HAMU_SmartJoin|supertip"
 HAMU_UITextPart2 = HAMU_L("İki tablodaki kayıtları müşteri kodu veya ürün kodu gibi ortak alanlarla yan yana eşleştirir. LEFT: ilk tablonun tüm satırları; INNER: yalnız eşleşenler; FULL: iki tablonun tüm kayıtları. İkinci tablodan alınacak sütunlar ayrıca seçilir. Kaynaklar değiştirilmez.", "Match two datasets using selected key columns. Check key types, spaces and duplicates before choosing the fields to include.")
 Case "r_HAMU_MergeDuplicates|label"
 HAMU_UITextPart2 = HAMU_L("Tekrarları Birleştir", "Merge Duplicate Records")
 Case "r_HAMU_MergeDuplicates|screentip"
 HAMU_UITextPart2 = HAMU_L("Tekrarları Birleştir", "Merge Duplicate Records")
 Case "r_HAMU_MergeDuplicates|supertip"
 HAMU_UITextPart2 = HAMU_L("Tek veya çoklu anahtar sütunlarla tekrar eden kayıtları tek satırda birleştirir. Seçtiğiniz sayısal sütunları toplar, diğer alanlarda ilk dolu değeri korur ve kaç kaydın birleştiğini ayrıca yazar.", "Merge records by one or more key columns, sum selected numeric fields, keep the first nonblank value in other fields and include a record count.")
 Case "r_HAMU_ConvertColumnTypes|label"
 HAMU_UITextPart2 = HAMU_L("Veri Türünü Düzelt", "Convert Types")
 Case "r_HAMU_ConvertColumnTypes|screentip"
 HAMU_UITextPart2 = HAMU_L("Sütun Veri Tipini Düzelt", "Fix Column Types")
 Case "r_HAMU_ConvertColumnTypes|supertip"
 HAMU_UITextPart2 = HAMU_L("Bir veya birden fazla sütunu metin, sayı, tarih, tam sayı veya yüzde tipine dönüştürür; dönüştürülemeyen hücreleri sayarak işlem sonucunu raporlar.", "Convert selected columns to the requested data types. Review mixed values and conversion errors.")
 Case "r_HAMU_SplitByValueToFiles|label"
 HAMU_UITextPart2 = HAMU_L("Dosyalara Böl", "Split to Files")
 Case "r_HAMU_SplitByValueToFiles|screentip"
 HAMU_UITextPart2 = HAMU_L("Değere Göre Ayrı Dosyalar Oluştur", "Split into Files")
 Case "r_HAMU_SplitByValueToFiles|supertip"
 HAMU_UITextPart2 = HAMU_L("Seçtiğiniz sütunun benzersiz değerlerine göre veriyi filtreler ve her değer için ayrı bir Excel dosyası oluşturur.", "Create separate files for distinct values in a selected column. Choose the destination folder and preserve the source data.")
 Case "r_HAMU_SplitURLParameters|label"
 HAMU_UITextPart2 = HAMU_L("URL Ayıkla", "Extract URL")
 Case "r_HAMU_SplitURLParameters|screentip"
 HAMU_UITextPart2 = HAMU_L("URL Parametrelerini Sütunlara Ayır", "Expand URL Parameters")
 Case "r_HAMU_SplitURLParameters|supertip"
 HAMU_UITextPart2 = HAMU_L("URL'lerde soru işaretinden sonra bulunan sorgu parametrelerini algılar ve her parametreyi ayrı bir sütun halinde yeni sayfaya çıkarır.", "Extract URL query parameters into separate columns for analysis.")
 Case "r_HAMU_PivotExport|label"
 HAMU_UITextPart2 = HAMU_L("Pivotu Böl", "Split Pivot")
 Case "r_HAMU_PivotExport|screentip"
 HAMU_UITextPart2 = HAMU_L("Pivot Filtrelerini Ayrı Sayfalara Aktar", "Export Pivot Filters")
 Case "r_HAMU_PivotExport|supertip"
 HAMU_UITextPart2 = HAMU_L("Seçili PivotTable'ın rapor filtresindeki değerleri sırayla uygular ve her filtre sonucu için ayrı çalışma sayfası oluşturur.", "Export separate PivotTable filter results to worksheets. Select a cell within the PivotTable first.")
 Case "r_HAMU_FilterRowsToNewTable|label"
 HAMU_UITextPart2 = HAMU_L("Koşulla Satır Filtrele", "Filter Rows")
 Case "r_HAMU_FilterRowsToNewTable|screentip"
 HAMU_UITextPart2 = HAMU_L("Koşulla Satır Filtrele", "Filter Rows")
 Case "r_HAMU_FilterRowsToNewTable|supertip"
 HAMU_UITextPart2 = HAMU_L("Koşula uyan satırları yeni sayfaya kopyalar; kaynak değişmez. Şehir=Bursa yalnız Bursa kayıtlarını, Satış>0 pozitif satışları, Ürün~TV içinde TV geçenleri seçer. Noktalı virgülle birleştirilen koşulların hepsi sağlanmalıdır.", "Create an output table containing rows that satisfy the chosen column condition. Source data remains unchanged.")
 Case "r_HAMU_SampleData|label"
 HAMU_UITextPart2 = HAMU_L("Veriden Örnek Al", "Sample Data")
 Case "r_HAMU_SampleData|screentip"
 HAMU_UITextPart2 = HAMU_L("Veriden Örnek Al", "Sample Data")
 Case "r_HAMU_SampleData|supertip"
 HAMU_UITextPart2 = HAMU_L("Tüm veri kümesinden rastgele N satır seçebilir veya belirlediğiniz grup sütunlarına göre her gruptan en fazla N rastgele kayıt alarak yeni bir örnek veri sayfası oluşturabilir.", "Create a sample of records using the requested sample size or selection options.")
 Case "r_HAMU_ShuffleRows|label"
 HAMU_UITextPart2 = HAMU_L("Satırları Karıştır", "Shuffle Rows")
 Case "r_HAMU_ShuffleRows|screentip"
 HAMU_UITextPart2 = HAMU_L("Satırları Karıştır", "Shuffle Rows")
 Case "r_HAMU_ShuffleRows|supertip"
 HAMU_UITextPart2 = HAMU_L("Kaynak veriyi değiştirmeden, başlık satırını koruyarak veri satırlarını rastgele sırada yeni bir çalışma sayfasına aktarır.", "Randomize the order of data rows while preserving values within each row.")
 Case "r_HAMU_CompareTwoLists|label"
 HAMU_UITextPart2 = HAMU_L("İki Listeyi Karşılaştır", "Compare Lists")
 Case "r_HAMU_CompareTwoLists|screentip"
 HAMU_UITextPart2 = HAMU_L("İki Listeyi Karşılaştır", "Compare Lists")
 Case "r_HAMU_CompareTwoLists|supertip"
 HAMU_UITextPart2 = HAMU_L("İki tek sütunlu listeyi karşılaştırır; yalnızca birinci listede, yalnızca ikinci listede veya her iki listede bulunan değerleri raporlar.", "Compare values in two lists and report matches and differences.")
 Case "g3|label"
 HAMU_UITextPart2 = HAMU_L("Metin ve Hücre", "Text and Cells")
 Case "quick_HAMU_CleanData|label"
 HAMU_UITextPart2 = HAMU_L(("Temizle" & vbLf & "Kırp"), ("Clean" & vbLf & "Trim"))
 Case "quick_HAMU_CleanData|screentip"
 HAMU_UITextPart2 = HAMU_L("Metin ve Boşluk Temizle", "Clean Text and Spaces")
 Case "quick_HAMU_CleanData|supertip"
 HAMU_UITextPart2 = HAMU_L("Seçili metinlerde görünmeyen NBSP karakterlerini, sekmeleri, satır sonlarını, sıfır genişlikli boşlukları ve art arda gelen gereksiz boşlukları temizler. Metnin başı/sonundaki boşlukları kaldırır; formülleri değiştirmez.", "Clean redundant spaces and invisible characters while preserving Turkish characters.")
 Case "quick_HAMU_TextCase|label"
 HAMU_UITextPart2 = HAMU_L("Harf Dönüştür", "Change Case")
 Case "quick_HAMU_TextCase|screentip"
 HAMU_UITextPart2 = HAMU_L("Harf Dönüştür", "Change Case")
 Case "quick_HAMU_TextCase|supertip"
 HAMU_UITextPart2 = HAMU_L("Seçili metinleri tamamı büyük, tamamı küçük veya baş harfleri büyük biçime dönüştürür.", "Choose uppercase, lowercase or title case from a simple selection dialog.")
 Case "r_HAMU_DeleteBlankCells|label"
 HAMU_UITextPart2 = HAMU_L("Boşlukları Sil", "Delete Blanks")
 Case "r_HAMU_DeleteBlankCells|screentip"
 HAMU_UITextPart2 = HAMU_L("Boşlukları Sil", "Delete Blanks")
 Case "r_HAMU_DeleteBlankCells|supertip"
 HAMU_UITextPart2 = HAMU_L("Seçili alandaki boş hücreleri silerek alttaki hücreleri yukarı kaydırır. Veri hizasını etkileyebileceği için tek sütunlu listelerde kullanılması önerilir.", "Delete empty cells and shift remaining cells upward. Check related columns before changing their alignment.")
 Case "m3|label"
 HAMU_UITextPart2 = HAMU_L("Düzenle", "Edit")
 Case "m3|screentip"
 HAMU_UITextPart2 = HAMU_L("Metin ve Hücre — Diğer Araçlar", "Text and Cell — Other Tools")
 Case "m3|supertip"
 HAMU_UITextPart2 = HAMU_L("Bu kategorideki diğer komutları ve açıklamalarını gösterir.", "Shows other commands and descriptions in this category.")
 Case "r_HAMU_AddTextAffixes|label"
 HAMU_UITextPart2 = HAMU_L("Başa / Sona Ekle", "Add Prefix or Suffix")
 Case "r_HAMU_AddTextAffixes|screentip"
 HAMU_UITextPart2 = HAMU_L("Başa / Sona Ekle", "Add Prefix or Suffix")
 Case "r_HAMU_AddTextAffixes|supertip"
 HAMU_UITextPart2 = HAMU_L("Seçili tek sütundaki metinlerin başına ve/veya sonuna metin ekler. Sonuç seçtiğiniz başlangıç hücresine yazılır; kaynak korunur. Boş hücreler boş kalır.", "Add text before or after values and write to a separate output range. Preserve the source and leading zeros.")
 Case "r_HAMU_LeadingZeros|label"
 HAMU_UITextPart2 = HAMU_L("Baştaki Sıfırlar", "Leading Zeros")
 End Select
End Function

Private Function HAMU_UITextPart3(ByVal key As String) As String
 Select Case key
 Case "r_HAMU_LeadingZeros|screentip"
 HAMU_UITextPart3 = HAMU_L("Baştaki Sıfırlar", "Leading Zeros")
 Case "r_HAMU_LeadingZeros|supertip"
 HAMU_UITextPart3 = HAMU_L("Kodları sabit uzunluğa baştan sıfır ekleyerek düzenler veya baştaki sıfırları kaldırır. Sonuç metin olarak seçtiğiniz hedefe yazılır; uzun kodlar ve kaynak korunur.", "Add or remove leading zeros from text codes. Preserve source values and write the result as text.")
 Case "r_HAMU_DeduplicateItems|label"
 HAMU_UITextPart3 = HAMU_L("Hücre İçi Tekrarlar", "Deduplicate Cell Items")
 Case "r_HAMU_DeduplicateItems|screentip"
 HAMU_UITextPart3 = HAMU_L("Hücre İçi Tekrarlar", "Deduplicate Cell Items")
 Case "r_HAMU_DeduplicateItems|supertip"
 HAMU_UITextPart3 = HAMU_L(("Virgül veya belirttiğiniz ayraçla ayrılmış hücre içi öğeleri tekilleştirir. Örneğin elma, armut, elma " & ChrW(8594) & " elma, armut. İlk sıra korunur; sonuç ayrı hedefe yazılır."), "Remove repeated items within each cell using the chosen separator, keeping their first occurrence.")
 Case "r_HAMU_ExtractTextPart|label"
 HAMU_UITextPart3 = HAMU_L("Metin Parçası Al", "Extract Text Part")
 Case "r_HAMU_ExtractTextPart|screentip"
 HAMU_UITextPart3 = HAMU_L("Metin Parçası Al", "Extract Text Part")
 Case "r_HAMU_ExtractTextPart|supertip"
 HAMU_UITextPart3 = HAMU_L("Soldan/sağdan belirli sayıda karakteri, bir işaretin öncesini/sonrasını veya iki işaret arasını ayıklar. İlk eşleşmeyi kullanır; eşleşme yoksa boş sonuç verir. Kaynak korunur.", "Extract characters from the left/right or text before, after or between specified markers.")
 Case "r_HAMU_ExtractDigits|label"
 HAMU_UITextPart3 = HAMU_L("Metinden Sayıları Ayıkla", "Extract Digits")
 Case "r_HAMU_ExtractDigits|screentip"
 HAMU_UITextPart3 = HAMU_L("Metinden Sayıları Ayıkla", "Extract Digits")
 Case "r_HAMU_ExtractDigits|supertip"
 HAMU_UITextPart3 = HAMU_L("Metin içindeki rakam karakterlerini ayıklar ve sonucu hedef alana aktarır.", "Extract numeric characters from selected text values.")
 Case "r_HAMU_MaskData|label"
 HAMU_UITextPart3 = HAMU_L("Veriyi Maskele", "Mask Data")
 Case "r_HAMU_MaskData|screentip"
 HAMU_UITextPart3 = HAMU_L("Veriyi Maskele", "Mask Data")
 Case "r_HAMU_MaskData|supertip"
 HAMU_UITextPart3 = HAMU_L("Seçili metinlerin belirlediğiniz kadar başlangıç ve bitiş karakterini koruyup ortadaki karakterleri yıldızla gizler.", "Mask selected text according to the available options. Keep an original copy when required.")
 Case "quick_HAMU_RegexReplace|label"
 HAMU_UITextPart3 = HAMU_L("Normal İfade ile Değiştir", "Replace by Regular Expression")
 Case "quick_HAMU_RegexReplace|screentip"
 HAMU_UITextPart3 = HAMU_L("Normal İfade ile Değiştir", "Replace by Regular Expression")
 Case "quick_HAMU_RegexReplace|supertip"
 HAMU_UITextPart3 = HAMU_L("Seçili metin hücrelerinde RegEx normal ifadeleri kullanarak toplu bul/değiştir işlemi yapar. Gruplama ve $1 gibi geri başvurular kullanılabilir.", "Replace matching text using a regular expression. Capture groups and references such as $1 can be used.")
 Case "r_HAMU_PatternFind|label"
 HAMU_UITextPart3 = HAMU_L("Normal İfade ile Bul", "Find by Regular Expression")
 Case "r_HAMU_PatternFind|screentip"
 HAMU_UITextPart3 = HAMU_L("Normal İfade ile Bul", "Find by Regular Expression")
 Case "r_HAMU_PatternFind|supertip"
 HAMU_UITextPart3 = HAMU_L("Seçili hücrelerde RegEx normal ifade desenine uyan metinleri bulur ve eşleşen hücreleri vurgular. Sayı, kod, telefon veya belirli metin kalıplarını tespit etmek için kullanılabilir.", "Find text matching a regular-expression pattern and highlight matching cells. Use a preset or a custom pattern.")
 Case "r_HAMU_StandardFormat|label"
 HAMU_UITextPart3 = HAMU_L("Standart Sayı ve Tarih Biçimi", "Number and Date Format")
 Case "r_HAMU_StandardFormat|screentip"
 HAMU_UITextPart3 = HAMU_L("Standart Sayı ve Tarih Biçimi", "Number and Date Format")
 Case "r_HAMU_StandardFormat|supertip"
 HAMU_UITextPart3 = HAMU_L("Gerçek sayıları ondalıksız tam sayı görünümüyle, tarih hücrelerini gg.aa.yyyy biçiminde gösterir. Hücre değerlerini yuvarlamaz veya değiştirmez.", "Apply integer number formatting or standard date formatting. Formatting controls display and does not automatically repair data types.")
 Case "r_HAMU_ConvertQuotedNumbers|label"
 HAMU_UITextPart3 = HAMU_L("Metin Olarak Saklanan Sayıları Düzelt", "Convert Text Numbers")
 Case "r_HAMU_ConvertQuotedNumbers|screentip"
 HAMU_UITextPart3 = HAMU_L("Metin Olarak Saklanan Sayıları Düzelt", "Convert Text Numbers")
 Case "r_HAMU_ConvertQuotedNumbers|supertip"
 HAMU_UITextPart3 = HAMU_L("Başında gizli tek tırnak bulunan veya Excel tarafından metin olarak saklanan sayısal değerleri gerçek sayıya dönüştürmeye çalışır.", "Convert suitable numbers stored as text to numeric values. Review decimal and thousands separators.")
 Case "r_HAMU_ConvertFormulasToValues|label"
 HAMU_UITextPart3 = HAMU_L("Formülleri Değere Dönüştür", "Formulas to Values")
 Case "r_HAMU_ConvertFormulasToValues|screentip"
 HAMU_UITextPart3 = HAMU_L("Formülleri Değere Dönüştür", "Formulas to Values")
 Case "r_HAMU_ConvertFormulasToValues|supertip"
 HAMU_UITextPart3 = HAMU_L("Seçili hücrelerdeki formülleri mevcut sonuçlarıyla değiştirir. İşlem sonrasında hücrelerde formül yerine sabit değer kalır.", "Replace selected formulas with their current calculated results.")
 Case "r_HAMU_UnmergeAndFill|label"
 HAMU_UITextPart3 = HAMU_L("Birleşimleri Çöz ve Doldur", "Unmerge and Fill")
 Case "r_HAMU_UnmergeAndFill|screentip"
 HAMU_UITextPart3 = HAMU_L("Birleşimleri Çöz ve Doldur", "Unmerge and Fill")
 Case "r_HAMU_UnmergeAndFill|supertip"
 HAMU_UITextPart3 = HAMU_L("Birleştirilmiş hücreleri çözer ve birleşik alanın sol üst değerini çözülen hücrelerin tamamına yazar.", "Unmerge cells and fill the resulting area with the original merged value.")
 Case "r_HAMU_FillDownBlanks|label"
 HAMU_UITextPart3 = HAMU_L("Boşlukları Aşağı Doldur", "Fill Down Blanks")
 Case "r_HAMU_FillDownBlanks|screentip"
 HAMU_UITextPart3 = HAMU_L("Boşlukları Aşağı Doldur", "Fill Down Blanks")
 Case "r_HAMU_FillDownBlanks|supertip"
 HAMU_UITextPart3 = HAMU_L("Seçili aralıktaki boş hücreleri aynı sütundaki en yakın üst dolu değerle doldurur. Pivot veya rapor çıktılarındaki grup etiketlerini satırlara yaymak için kullanışlıdır.", "Fill blank cells with the nearest nonblank value above in the same column. Useful for repeated group labels in reports.")
 Case "r_HAMU_FillBlanksZero|label"
 HAMU_UITextPart3 = HAMU_L("Boşlukları Sıfırla Doldur", "Fill Blanks with Zero")
 Case "r_HAMU_FillBlanksZero|screentip"
 HAMU_UITextPart3 = HAMU_L("Boşlukları Sıfırla Doldur", "Fill Blanks with Zero")
 Case "r_HAMU_FillBlanksZero|supertip"
 HAMU_UITextPart3 = HAMU_L("Seçili aralıktaki gerçekten boş hücrelere 0 yazar. Dolu hücreleri ve boş metin döndüren formülleri korur.", "Write zero only to genuinely empty cells. Preserve populated cells and formulas returning empty text.")
 Case "g4|label"
 HAMU_UITextPart3 = HAMU_L("Tablo", "Tables")
 Case "quick_HAMU_AutoTable|label"
 HAMU_UITextPart3 = HAMU_L(("Hızlı" & vbLf & "Tablo"), ("Quick" & vbLf & "Table"))
 Case "quick_HAMU_AutoTable|screentip"
 HAMU_UITextPart3 = HAMU_L("Otomatik Tablo Oluştur", "Create Table")
 Case "quick_HAMU_AutoTable|supertip"
 HAMU_UITextPart3 = HAMU_L("Tek hücre seçildiyse çevresindeki kesintisiz veri alanını, aksi halde seçilen alanı başlıklı Excel tablosuna dönüştürür. HAMU stili ve filtreleri uygular.", "Create an Excel table from a header/data range and apply a consistent table style. A single selected cell uses its current data region.")
 Case "quick_HAMU_AutoFit|label"
 HAMU_UITextPart3 = HAMU_L("Sığdır", "Fit Height")
 Case "quick_HAMU_AutoFit|screentip"
 HAMU_UITextPart3 = HAMU_L("Sığdır", "Fit Height")
 Case "quick_HAMU_AutoFit|supertip"
 HAMU_UITextPart3 = HAMU_L("Seçili alanın satır yüksekliğini metne göre sığdırır. Sütun genişliği değişmez.", "Fit row heights to selected content without changing column widths.")
 End Select
End Function

Private Function HAMU_UITextPart4(ByVal key As String) As String
 Select Case key
 Case "quick_HAMU_ResizeTable|label"
 HAMU_UITextPart4 = HAMU_L("Tabloyu Genişlet", "Expand Table")
 Case "quick_HAMU_ResizeTable|screentip"
 HAMU_UITextPart4 = HAMU_L("Tabloyu Genişlet", "Expand Table")
 Case "quick_HAMU_ResizeTable|supertip"
 HAMU_UITextPart4 = HAMU_L("Tablonun altındaki bitişik verileri tabloya dahil eder.", "Include adjacent data rows below the table.")
 Case "m4|label"
 HAMU_UITextPart4 = HAMU_L("Tablo Araçları", "Table Tools")
 Case "m4|screentip"
 HAMU_UITextPart4 = HAMU_L("Tablo ve Görünüm — Diğer Araçlar", "Table and View — Other Vehicles")
 Case "m4|supertip"
 HAMU_UITextPart4 = HAMU_L("Bu kategorideki diğer komutları ve açıklamalarını gösterir.", "Shows other commands and descriptions in this category.")
 Case "r_HAMU_SelectColumnsToNewTable|label"
 HAMU_UITextPart4 = HAMU_L("Sütun Kopyala", "Copy Columns")
 Case "r_HAMU_SelectColumnsToNewTable|screentip"
 HAMU_UITextPart4 = HAMU_L("Sütun Kopyala", "Copy Columns")
 Case "r_HAMU_SelectColumnsToNewTable|supertip"
 HAMU_UITextPart4 = HAMU_L("Başlık ve veri alanından sütunları adlarıyla veya seçim içindeki 1,2,3 gibi numaralarla seçer. Seçilen sütunlar ile yeni bir tablo oluşturulur.", "Select columns by header names or positions such as 1,2,3 within the selected data. Whole-column sources are trimmed to populated rows.")
 Case "g5|label"
 HAMU_UITextPart4 = HAMU_L("Tarih", "Dates")
 Case "quick_HAMU_DatePicker|label"
 HAMU_UITextPart4 = HAMU_L("Tarih Seç", "Pick Date")
 Case "quick_HAMU_DatePicker|screentip"
 HAMU_UITextPart4 = HAMU_L("Tarih Seçici", "Date Picker")
 Case "quick_HAMU_DatePicker|supertip"
 HAMU_UITextPart4 = HAMU_L("HAMU takvim penceresini açar. Aktif hücrede tarih varsa takvim o tarihle açılır; seçtiğiniz tarih hücreye gerçek Excel tarihi olarak dd.mm.yyyy biçiminde yazılır.", "Choose a date and write it to the previously selected active cell. Select month/year and see the distance from today.")
 Case "quick_HAMU_FixDates|label"
 HAMU_UITextPart4 = HAMU_L("Tarihleri Düzelt", "Fix Dates")
 Case "quick_HAMU_FixDates|screentip"
 HAMU_UITextPart4 = HAMU_L("Tarihleri Düzelt", "Fix Dates")
 Case "quick_HAMU_FixDates|supertip"
 HAMU_UITextPart4 = HAMU_L("03/10/2026, 2026-10-03, 20261003, 3 Ekim 2026 ve benzeri farklı tarih gösterimlerini gerçek Excel tarihine dönüştürür ve dd.mm.yyyy biçiminde standardize eder.", "Normalize suitable date values and review values that cannot be converted.")
 Case "quick_HAMU_CheckDates|label"
 HAMU_UITextPart4 = HAMU_L("Tarih Kontrolü", "Check Dates")
 Case "quick_HAMU_CheckDates|screentip"
 HAMU_UITextPart4 = HAMU_L("Tarih Kontrolü", "Check Dates")
 Case "quick_HAMU_CheckDates|supertip"
 HAMU_UITextPart4 = HAMU_L("Seçili hücrelerdeki değerlerin geçerli bir tarih olup olmadığını kontrol eder. Geçerli tarihleri ve tanımlanamayan değerleri farklı renklerle işaretler.", "Inspect date values for invalid or inconsistent entries.")
 Case "m5|label"
 HAMU_UITextPart4 = HAMU_L("Tarih Araçları", "Date Tools")
 Case "m5|screentip"
 HAMU_UITextPart4 = HAMU_L("Tarih Araçları — Diğer Araçlar", "Date Tools — Other Vehicles")
 Case "m5|supertip"
 HAMU_UITextPart4 = HAMU_L("Bu kategorideki diğer komutları ve açıklamalarını gösterir.", "Shows other commands and descriptions in this category.")
 Case "r_HAMU_DateBuilder|label"
 HAMU_UITextPart4 = HAMU_L("Tarih Oluştur", "Build Dates")
 Case "r_HAMU_DateBuilder|screentip"
 HAMU_UITextPart4 = HAMU_L("Tarih Oluştur", "Build Dates")
 Case "r_HAMU_DateBuilder|supertip"
 HAMU_UITextPart4 = HAMU_L("Ayrı sütunlarda bulunan yıl, ay ve gün değerlerini birleştirerek gerçek Excel tarihleri oluşturur. Ay alanında sayı veya Türkçe ay adı kullanılabilir.", "Create dates from separate year, month and day fields.")
 Case "r_HAMU_DateRangeList|label"
 HAMU_UITextPart4 = HAMU_L("Tarih Listesi", "Date List")
 Case "r_HAMU_DateRangeList|screentip"
 HAMU_UITextPart4 = HAMU_L("Tarih Listesi", "Date List")
 Case "r_HAMU_DateRangeList|supertip"
 HAMU_UITextPart4 = HAMU_L("Başlangıç ve bitiş dahil tarih listesi oluşturur. Tüm günler veya yalnız hafta içi seçilir; hafta içi seçeneği resmî tatilleri hesaplamaz. Sonuç seçtiğiniz başlangıç hücresine yazılır.", "Generate dates between a start and end date, optionally excluding weekends. Choose the output start cell.")
 Case "g6|label"
 HAMU_UITextPart4 = HAMU_L("Sayfa ve Görsel", "Sheet and Images")
 Case "quick_HAMU_SheetIndex|label"
 HAMU_UITextPart4 = HAMU_L("Sayfa İndeksi", "Sheet Index")
 Case "quick_HAMU_SheetIndex|screentip"
 HAMU_UITextPart4 = HAMU_L("Sayfa İndeksi", "Sheet Index")
 Case "quick_HAMU_SheetIndex|supertip"
 HAMU_UITextPart4 = HAMU_L("Görünür çalışma sayfalarını listeleyen ve her sayfaya tıklanabilir bağlantı veren bir İçindekiler sayfası oluşturur.", "Create a worksheet index with hyperlinks to the workbook's sheets.")
 Case "quick_HAMU_DynamicValidation|label"
 HAMU_UITextPart4 = HAMU_L("Dinamik Veri Doğrulama Listesi", "Dynamic Validation List")
 Case "quick_HAMU_DynamicValidation|screentip"
 HAMU_UITextPart4 = HAMU_L("Dinamik Veri Doğrulama Listesi", "Dynamic Validation List")
 Case "quick_HAMU_DynamicValidation|supertip"
 HAMU_UITextPart4 = HAMU_L("Seçili hücrelere liste tipi veri doğrulama uygular. Kaynak olarak adlandırılmış aralık, Tablo[Sütun], doğrudan aralık veya A1# biçimindeki dinamik taşan dizi kullanılabilir. Bir adlandırma A1# kaynağına bağlıysa taşan aralık büyüdükçe doğrulama listesi de dinamik kalır.", "Create list validation using a named range, Table[Column], direct range or spilled range such as A1#. Table/spill sources can expand dynamically.")
 Case "quick_HAMU_SelectionToPNG|label"
 HAMU_UITextPart4 = HAMU_L("Resim Kaydet", "Save Image")
 Case "quick_HAMU_SelectionToPNG|screentip"
 HAMU_UITextPart4 = HAMU_L("Resim Kaydet", "Save Image")
 Case "quick_HAMU_SelectionToPNG|supertip"
 HAMU_UITextPart4 = HAMU_L("Seçili alanı PNG, JPG veya GIF olarak kaydeder. Çok büyük görsel alanlarını durdurur ve geçici grafiği temizler.", "Save the selected range as PNG, JPG or GIF. Reject oversized images and clean up the temporary chart.")
 Case "m6|label"
 HAMU_UITextPart4 = HAMU_L("Sayfa Araçları", "Sheet Tools")
 Case "m6|screentip"
 HAMU_UITextPart4 = HAMU_L("Sayfa ve Görsel — Diğer Araçlar", "Page and Visual — Other Tools")
 Case "m6|supertip"
 HAMU_UITextPart4 = HAMU_L("Bu kategorideki diğer komutları ve açıklamalarını gösterir.", "Shows other commands and descriptions in this category.")
 Case "r_HAMU_SortSheets|label"
 HAMU_UITextPart4 = HAMU_L("Sayfaları Alfabetik Sırala", "Sort Sheets")
 Case "r_HAMU_SortSheets|screentip"
 HAMU_UITextPart4 = HAMU_L("Sayfaları Alfabetik Sırala", "Sort Sheets")
 Case "r_HAMU_SortSheets|supertip"
 HAMU_UITextPart4 = HAMU_L("Çalışma kitabındaki çalışma sayfalarını adlarına göre alfabetik olarak yeniden sıralar.", "Sort worksheet tabs alphabetically.")
 Case "r_HAMU_ShowHidden|label"
 HAMU_UITextPart4 = HAMU_L("Tüm Gizlileri Göster", "Show Hidden Items")
 End Select
End Function

Private Function HAMU_UITextPart5(ByVal key As String) As String
 Select Case key
 Case "r_HAMU_ShowHidden|screentip"
 HAMU_UITextPart5 = HAMU_L("Tüm Gizlileri Göster", "Show Hidden Items")
 Case "r_HAMU_ShowHidden|supertip"
 HAMU_UITextPart5 = HAMU_L("Aktif çalışma kitabındaki gizli çalışma sayfalarını, satırları ve sütunları görünür hale getirir.", "Reveal hidden sheets, rows or columns as supported by the command.")
 Case "r_HAMU_InsertPhotosFromFolder|label"
 HAMU_UITextPart5 = HAMU_L(("Toplu" & vbLf & "Fotoğraf"), ("Batch" & vbLf & "Photos"))
 Case "r_HAMU_InsertPhotosFromFolder|screentip"
 HAMU_UITextPart5 = HAMU_L("Klasörden Toplu Fotoğraf Ekle", "Insert Folder Images")
 Case "r_HAMU_InsertPhotosFromFolder|supertip"
 HAMU_UITextPart5 = HAMU_L("Seçtiğiniz klasördeki görselleri aktif hücreden başlayarak satırlara ekler ve yan hücreye dosya adını yazar.", "Insert supported images from a folder into cells or over cells. In-cell conversion depends on Excel version support; over-cell images are constrained to cell bounds.")
 Case "r_HAMU_DeleteDropdowns|label"
 HAMU_UITextPart5 = HAMU_L("Eski Dropdown Nesnelerini Sil", "Remove Old Dropdown Objects")
 Case "r_HAMU_DeleteDropdowns|screentip"
 HAMU_UITextPart5 = HAMU_L("Eski Dropdown Nesnelerini Sil", "Remove Old Dropdown Objects")
 Case "r_HAMU_DeleteDropdowns|supertip"
 HAMU_UITextPart5 = HAMU_L("Aktif sayfadaki eski Form Denetimi ve ActiveX açılır kutu nesnelerini siler. Hücrelerdeki veri doğrulama listeleri korunur.", "Remove old worksheet dropdown objects while preserving cell data-validation lists.")
 Case "r_HAMU_RowStriping|label"
 HAMU_UITextPart5 = HAMU_L("Satırları Renklendir", "Stripe Rows")
 Case "r_HAMU_RowStriping|screentip"
 HAMU_UITextPart5 = HAMU_L("Satırları Renklendir", "Stripe Rows")
 Case "r_HAMU_RowStriping|supertip"
 HAMU_UITextPart5 = HAMU_L("Seçili aralıkta dönüşümlü satır renklendirmesi uygular. Çalıştırıldığında mavi, yeşil, sarı, turuncu, mor, gri ve pembe tonlarından birini seçebilirsiniz.", "Apply alternating row colors to the selected area.")
 Case "r_HAMU_CreateQRCode|label"
 HAMU_UITextPart5 = HAMU_L(("Karekod" & vbLf & "Barkod"), ("QR Code" & vbLf & "Barcode"))
 Case "r_HAMU_CreateQRCode|screentip"
 HAMU_UITextPart5 = HAMU_L("Karekod ve Barkod", "QR and Barcode")
 Case "r_HAMU_CreateQRCode|supertip"
 HAMU_UITextPart5 = HAMU_L("Tek ekranda Karekod veya Barkod seçin; tür değişince model listesi güncellenir. Metin, bağlantı, e-posta ve telefon içeriği; tek giriş veya tek sütundan toplu üretim. Kodlar verinin sağındaki boş hücreye eklenir. İnternet gerekir; en fazla 250 kod.", "Select QR Code or Barcode in one form; model choices follow the type. Supports text, links, email, phone, single values or a batch column. Batch codes go in empty cells to the right. Internet required; at most 250 codes.")
 Case "g7|label"
 HAMU_UITextPart5 = HAMU_L("Denetim", "Audit")
 Case "quick_HAMU_HighlightDuplicates|label"
 HAMU_UITextPart5 = HAMU_L("Tekrar Edenleri Vurgula", "Highlight Duplicates")
 Case "quick_HAMU_HighlightDuplicates|screentip"
 HAMU_UITextPart5 = HAMU_L("Tekrar Edenleri Vurgula", "Highlight Duplicates")
 Case "quick_HAMU_HighlightDuplicates|supertip"
 HAMU_UITextPart5 = HAMU_L("Seçili alanda birden fazla kez bulunan değerleri tespit eder ve tekrar eden hücreleri renklendirir.", "Highlight repeated values within the selected area.")
 Case "quick_HAMU_CompareSheets|label"
 HAMU_UITextPart5 = HAMU_L("Sayfaları Karşılaştır", "Compare Sheets")
 Case "quick_HAMU_CompareSheets|screentip"
 HAMU_UITextPart5 = HAMU_L("Sayfaları Karşılaştır", "Compare Sheets")
 Case "quick_HAMU_CompareSheets|supertip"
 HAMU_UITextPart5 = HAMU_L("İki çalışma sayfasını belirlediğiniz anahtar sütuna göre karşılaştırır ve eklenen, silinen veya değişen kayıtları ayrı bir raporda gösterir.", "Compare worksheets and identify differences.")
 Case "quick_HAMU_MarkProblemCells|label"
 HAMU_UITextPart5 = HAMU_L(("Sorunlu" & vbLf & "Hücreler"), ("Problem" & vbLf & "Cells"))
 Case "quick_HAMU_MarkProblemCells|screentip"
 HAMU_UITextPart5 = HAMU_L("Sorunlu Hücreleri İşaretle", "Mark Problem Cells")
 Case "quick_HAMU_MarkProblemCells|supertip"
 HAMU_UITextPart5 = HAMU_L("Sayı gibi görünen metinleri ve eşittir işaretiyle başlayan ancak gerçek formül olmayan metinleri bularak hücreleri farklı renklerle işaretler.", "Highlight cells with the problems selected in the dialog.")
 Case "m7|label"
 HAMU_UITextPart5 = HAMU_L(("Denetim" & vbLf & "Araçları"), ("Audit" & vbLf & "Tools"))
 Case "m7|screentip"
 HAMU_UITextPart5 = HAMU_L("Kontrol ve Denetim — Diğer Araçlar", "Control and Control — Other Vehicles")
 Case "m7|supertip"
 HAMU_UITextPart5 = HAMU_L("Bu kategorideki diğer komutları ve açıklamalarını gösterir.", "Shows other commands and descriptions in this category.")
 Case "r_HAMU_HighlightByCriteria|label"
 HAMU_UITextPart5 = HAMU_L("Hücre Türüne Göre Vurgula", "Highlight Cell Types")
 Case "r_HAMU_HighlightByCriteria|screentip"
 HAMU_UITextPart5 = HAMU_L("Hücre Türüne Göre Vurgula", "Highlight Cell Types")
 Case "r_HAMU_HighlightByCriteria|supertip"
 HAMU_UITextPart5 = HAMU_L("Seçili alanda sayı, metin, hata veya boş hücreleri seçtiğiniz kritere göre vurgular.", "Highlight cells matching a selected type or criterion.")
 Case "r_HAMU_AddDuplicateCount|label"
 HAMU_UITextPart5 = HAMU_L("Tekrar Sayısını Yaz", "Write Duplicate Counts")
 Case "r_HAMU_AddDuplicateCount|screentip"
 HAMU_UITextPart5 = HAMU_L("Tekrar Sayısını Yaz", "Write Duplicate Counts")
 Case "r_HAMU_AddDuplicateCount|supertip"
 HAMU_UITextPart5 = HAMU_L("Giriş listesindeki her değerin kaç kez geçtiğini hesaplayıp ayrı seçtiğiniz çıkış alanına yazar. Tek çıkış hücresi seçerseniz sonuç aşağı doğru genişler.", "Count occurrences of each input value and write the counts to a separately selected output range. A single output cell defines the starting point.")
 Case "r_HAMU_ListLinks|label"
 HAMU_UITextPart5 = HAMU_L("Harici Bağlantıları Listele", "List External Links")
 Case "r_HAMU_ListLinks|screentip"
 HAMU_UITextPart5 = HAMU_L("Harici Bağlantıları Listele", "List External Links")
 Case "r_HAMU_ListLinks|supertip"
 HAMU_UITextPart5 = HAMU_L("Çalışma kitabındaki Excel türü dış bağlantıları yeni bir sayfada listeler.", "List external Excel workbook links on a new worksheet.")
 Case "r_HAMU_BreakLinks|label"
 HAMU_UITextPart5 = HAMU_L("Harici Bağlantıları Kır", "Break External Links")
 Case "r_HAMU_BreakLinks|screentip"
 HAMU_UITextPart5 = HAMU_L("Harici Bağlantıları Kır", "Break External Links")
 Case "r_HAMU_BreakLinks|supertip"
 HAMU_UITextPart5 = HAMU_L("Çalışma kitabındaki harici Excel bağlantılarını kırar. Bağlantılı formüller mevcut değerlerine dönüştürülebileceği için işlemden önce yedek alınması önerilir.", "Break external Excel links. Linked formulas may become values; back up the workbook before proceeding.")
 Case "r_HAMU_CountByColor|label"
 HAMU_UITextPart5 = HAMU_L("Renge Göre Say", "Count by Color")
 Case "r_HAMU_CountByColor|screentip"
 HAMU_UITextPart5 = HAMU_L("Renge Göre Say", "Count by Color")
 Case "r_HAMU_CountByColor|supertip"
 HAMU_UITextPart5 = HAMU_L("Seçilen aralıkta örnek hücrenin görüntülenen dolgu rengiyle eşleşen hücreleri sayar; koşullu biçimlendirme dahildir.", "Count cells matching the sample cell's displayed fill color, including conditional formatting.")
 Case "r_HAMU_SumByColor|label"
 HAMU_UITextPart5 = HAMU_L("Renge Göre Topla", "Sum by Color")
 Case "r_HAMU_SumByColor|screentip"
 HAMU_UITextPart5 = HAMU_L("Renge Göre Topla", "Sum by Color")
 Case "r_HAMU_SumByColor|supertip"
 HAMU_UITextPart5 = HAMU_L("Örnek hücrenin görüntülenen dolgu rengiyle eşleşen sayısal değerleri toplar; metin, boş ve hata değerlerini toplama katmaz.", "Sum numeric values matching the sample cell's displayed fill color, including conditional formatting. Ignore blank, text and error values.")
 End Select
End Function

Private Function HAMU_UITextPart6(ByVal key As String) As String
 Select Case key
 Case "r_HAMU_FormulaAudit|label"
 HAMU_UITextPart6 = HAMU_L("Formül Denetimi", "Formula Audit")
 Case "r_HAMU_FormulaAudit|screentip"
 HAMU_UITextPart6 = HAMU_L("Formül Denetimi", "Formula Audit")
 Case "r_HAMU_FormulaAudit|supertip"
 HAMU_UITextPart6 = HAMU_L("Kitaptaki formülleri kaynak hücre bağlantısı, formül metni, görüntülenen sonuç, hata durumu ve dış kitap başvurusu bilgisiyle yeni sayfada listeler. Kaynak formüller değiştirilmez. Dış başvuru tespiti doğrudan XLS/CSV kitap başvurularını tarar.", "List formulas, source-cell links, displayed results, errors and direct external workbook references on a new sheet. Source formulas remain unchanged.")
 Case "r_HAMU_SelectSpecialCells|label"
 HAMU_UITextPart6 = HAMU_L("Özel Hücreleri Seç", "Select Special Cells")
 Case "r_HAMU_SelectSpecialCells|screentip"
 HAMU_UITextPart6 = HAMU_L("Özel Hücreleri Seç", "Select Special Cells")
 Case "r_HAMU_SelectSpecialCells|supertip"
 HAMU_UITextPart6 = HAMU_L("Mevcut seçimdeki formül, sabit değer, hata, boş veya görünür hücreleri seçer. Değer ve biçimler değiştirilmez; boş metin döndüren formüller gerçek boş sayılmaz.", "Select formulas, constants, errors, blanks or visible cells within the current area. Do not change values or formatting.")
 Case "r_HAMU_ListDefinedNames|label"
 HAMU_UITextPart6 = HAMU_L("Tanımlı Adları Listele", "List Defined Names")
 Case "r_HAMU_ListDefinedNames|screentip"
 HAMU_UITextPart6 = HAMU_L("Tanımlı Adları Listele", "List Defined Names")
 Case "r_HAMU_ListDefinedNames|supertip"
 HAMU_UITextPart6 = HAMU_L("Kitap ve sayfa kapsamındaki tanımlı adları, başvurularını, görünürlüklerini ve geçersiz başvuru durumlarını yeni sayfada listeler. Adları değiştirmez veya silmez.", "List workbook/worksheet names, references, visibility and broken references on a new sheet without modifying the names.")
 Case "g8|label"
 HAMU_UITextPart6 = HAMU_L("Bilgi", "Information")
 Case "quick_HAMU_Help|label"
 HAMU_UITextPart6 = HAMU_L("Yardım", "Help")
 Case "quick_HAMU_Help|screentip"
 HAMU_UITextPart6 = HAMU_L("Yardım", "Help")
 Case "quick_HAMU_Help|supertip"
 HAMU_UITextPart6 = HAMU_L("HAMU Tools içindeki komutların kullanım amacı ve temel kullanım bilgilerini gösterir.", "Open the usage guide with workflows, selection/output explanations, limitations and the complete tool catalog.")
 Case "quick_HAMU_About|label"
 HAMU_UITextPart6 = HAMU_L("Hakkında", "About")
 Case "quick_HAMU_About|screentip"
 HAMU_UITextPart6 = HAMU_L("Hakkında", "About")
 Case "quick_HAMU_About|supertip"
 HAMU_UITextPart6 = HAMU_L("HAMU Tools V1 sürümü, HAMU markası ve huseyinavniuzun.com bilgileri ile eklentinin kapsamını gösterir.", "Show HAMU Tools V1, Hüseyin Avni UZUN and the website, GitHub and email links.")
 Case "m8|label"
 HAMU_UITextPart6 = HAMU_L("Bilgi", "Information")
 Case "m8|screentip"
 HAMU_UITextPart6 = HAMU_L("Yardım ve Hakkında — Diğer Araçlar", "Help and About — Other Vehicles")
 Case "m8|supertip"
 HAMU_UITextPart6 = HAMU_L("Bu kategorideki diğer komutları ve açıklamalarını gösterir.", "Shows other commands and descriptions in this category.")
 Case "r_HAMU_UDFHelp|label"
 HAMU_UITextPart6 = HAMU_L("Çalışma Sayfası Fonksiyonları", "Worksheet Functions")
 Case "r_HAMU_UDFHelp|screentip"
 HAMU_UITextPart6 = HAMU_L("Çalışma Sayfası Fonksiyonları", "Worksheet Functions")
 Case "r_HAMU_UDFHelp|supertip"
 HAMU_UITextPart6 = HAMU_L("HAMU çalışma sayfası fonksiyonlarını ve formül örneklerini gösterir.", "Show the eight HAMU worksheet functions with examples. Function names remain unchanged when switching language.")
 Case "r_UDF_URLDecode|label"
 HAMU_UITextPart6 = HAMU_L("HAMU_URLCoz", "HAMU_URLCoz")
 Case "r_UDF_URLDecode|screentip"
 HAMU_UITextPart6 = HAMU_L("HAMU_URLCoz", "HAMU_URLCoz")
 Case "r_UDF_URLDecode|supertip"
 HAMU_UITextPart6 = HAMU_L("=HAMU_URLCoz(A1) | UTF-8 URL karakterlerini çözer. Geçersiz yüzde dizisini korur.", "=HAMU_URLCoz(A1) | Decode UTF-8 URL characters; preserve invalid percent sequences.")
 Case "r_UDF_RenklileriTopla|label"
 HAMU_UITextPart6 = HAMU_L("HAMU_RengeGoreTopla", "HAMU_RengeGoreTopla")
 Case "r_UDF_RenklileriTopla|screentip"
 HAMU_UITextPart6 = HAMU_L("HAMU_RengeGoreTopla", "HAMU_RengeGoreTopla")
 Case "r_UDF_RenklileriTopla|supertip"
 HAMU_UITextPart6 = HAMU_L("=HAMU_RengeGoreTopla(A1:A20;B1) | B1 ile aynı dolgu rengine sahip sayıları toplar. Biçim değişimi sonrası F9 kullanın.", "=HAMU_RengeGoreTopla(A1:A20;B1) | Sum numbers with the same direct fill color as B1. Recalculate with F9 after changing formatting.")
 Case "r_UDF_RengeGoreSay|label"
 HAMU_UITextPart6 = HAMU_L("HAMU_RengeGoreSay", "HAMU_RengeGoreSay")
 Case "r_UDF_RengeGoreSay|screentip"
 HAMU_UITextPart6 = HAMU_L("HAMU_RengeGoreSay", "HAMU_RengeGoreSay")
 Case "r_UDF_RengeGoreSay|supertip"
 HAMU_UITextPart6 = HAMU_L("=HAMU_RengeGoreSay(A1:A20;B1) | B1 ile aynı dolgu rengine sahip hücreleri sayar. Biçim değişimi sonrası F9 kullanın.", "=HAMU_RengeGoreSay(A1:A20;B1) | Count cells with the same direct fill color as B1. Recalculate with F9 after changing formatting.")
 Case "r_UDF_YerelDosyaYolu|label"
 HAMU_UITextPart6 = HAMU_L("HAMU_YerelDosyaYolu", "HAMU_YerelDosyaYolu")
 Case "r_UDF_YerelDosyaYolu|screentip"
 HAMU_UITextPart6 = HAMU_L("HAMU_YerelDosyaYolu", "HAMU_YerelDosyaYolu")
 Case "r_UDF_YerelDosyaYolu|supertip"
 HAMU_UITextPart6 = HAMU_L("=HAMU_YerelDosyaYolu(A1) | Kişisel OneDrive URL yolunu yerel OneDrive ortam değişkeniyle eşler; diğer URL türlerini korur.", "=HAMU_YerelDosyaYolu(A1) | Map a personal OneDrive URL to a local path using the local OneDrive environment variable; preserve other URL types.")
 Case "r_UDF_GizliLink|label"
 HAMU_UITextPart6 = HAMU_L("HAMU_GizliLink", "HAMU_GizliLink")
 Case "r_UDF_GizliLink|screentip"
 HAMU_UITextPart6 = HAMU_L("HAMU_GizliLink", "HAMU_GizliLink")
 Case "r_UDF_GizliLink|supertip"
 HAMU_UITextPart6 = HAMU_L("=HAMU_GizliLink(A1) | Hücrenin ilk köprü adresini getirir.", "=HAMU_GizliLink(A1) | Return the first hyperlink address in the cell.")
 Case "r_UDF_LinkParamDegeri|label"
 HAMU_UITextPart6 = HAMU_L("HAMU_LinkParamDegeri", "HAMU_LinkParamDegeri")
 Case "r_UDF_LinkParamDegeri|screentip"
 HAMU_UITextPart6 = HAMU_L("HAMU_LinkParamDegeri", "HAMU_LinkParamDegeri")
 Case "r_UDF_LinkParamDegeri|supertip"
 HAMU_UITextPart6 = HAMU_L("=HAMU_LinkParamDegeri(A1;""id"") | URL içindeki sorgu parametresinin değerini getirir.", "=HAMU_LinkParamDegeri(A1;""id"") | Return the value of the specified URL query parameter.")
 Case "r_UDF_SayidanMetine|label"
 HAMU_UITextPart6 = HAMU_L("HAMU_SayidanMetine", "HAMU_SayidanMetine")
 Case "r_UDF_SayidanMetine|screentip"
 HAMU_UITextPart6 = HAMU_L("HAMU_SayidanMetine", "HAMU_SayidanMetine")
 Case "r_UDF_SayidanMetine|supertip"
 HAMU_UITextPart6 = HAMU_L("=HAMU_SayidanMetine(123) | Türkçe: yüz yirmi üç. 15 basamaktan küçük mutlak değerler; ondalık kısmı tek tek okur.", "=HAMU_SayidanMetine(123) | Spell numbers in Turkish. Supports absolute values below 15 digits and reads decimal digits individually. The output language stays Turkish.")
 Case "r_UDF_TablodanVeriGetir|label"
 HAMU_UITextPart6 = HAMU_L("HAMU_TablodanVeriGetir", "HAMU_TablodanVeriGetir")
 Case "r_UDF_TablodanVeriGetir|screentip"
 HAMU_UITextPart6 = HAMU_L("HAMU_TablodanVeriGetir", "HAMU_TablodanVeriGetir")
 End Select
End Function

Private Function HAMU_UITextPart7(ByVal key As String) As String
 Select Case key
 Case "r_UDF_TablodanVeriGetir|supertip"
 HAMU_UITextPart7 = HAMU_L("=HAMU_TablodanVeriGetir(A1;""Tablo1"";""Kod"";""Ad"") | Formülün bulunduğu kitapta tablodan ilk eşleşen değeri getirir.", "=HAMU_TablodanVeriGetir(A1;""Tablo1"";""Kod"";""Ad"") | Return the first matching value from a table in the workbook containing the formula.")
 Case "bsHAMU|label"
 HAMU_UITextPart7 = HAMU_L("HAMU Tools", "HAMU Tools")
 Case "guide0|label"
 HAMU_UITextPart7 = HAMU_L("01  Nereden başlamalıyım?", "01  Where Do I Start?")
 Case "guide0|helperText"
 HAMU_UITextPart7 = HAMU_L("Kaynak veriyi seçin, uygun aracı açın ve açıklamayı okuyun. İlk kez kullanıyorsanız dört adımlı başlangıç rehberini inceleyin.", "Select source data, open the appropriate tool and read the description. Check the four-step start guide if you’re using the first time.")
 Case "guideButton0|label"
 HAMU_UITextPart7 = HAMU_L("Rehberi aç", "Open Guide")
 Case "guide1|label"
 HAMU_UITextPart7 = HAMU_L("02  Veriyi temizleyin", "02  Clean Your Data")
 Case "guide1|helperText"
 HAMU_UITextPart7 = HAMU_L("Başlıkları düzenleyin, boşlukları ve tekrarları inceleyin. Kodların başındaki sıfırları koruyun; metin ve sayı türlerini kontrol edin.", "Edit headers, view blanks and repetitions. Protect zeros at the beginning of the codes; check text and number types.")
 Case "guideButton1|label"
 HAMU_UITextPart7 = HAMU_L("Rehberi aç", "Open Guide")
 Case "guide2|label"
 HAMU_UITextPart7 = HAMU_L("03  Listeden anlamlı sonuçlara", "03  Turn Lists into Results")
 Case "guide2|helperText"
 HAMU_UITextPart7 = HAMU_L("Aynı yapılı listeleri birleştirin, ortak anahtarla eşleştirin; Grupla, Pivot ve Unpivot ile veri yapısını dönüştürün.", "Combine the same structured lists, match with common key; Convert data structure with group, Pivot and Unpivot.")
 Case "guideButton2|label"
 HAMU_UITextPart7 = HAMU_L("Rehberi aç", "Open Guide")
 Case "guide3|label"
 HAMU_UITextPart7 = HAMU_L("04  Dosyalarınızı yönetin", "04  Manage Your Files")
 Case "guide3|helperText"
 HAMU_UITextPart7 = HAMU_L("Önce tarih ve saatli bir yedek alın. Klasör envanteri, dosyalardan veri toplama ve şablon aktarımı için uygun akışı seçin.", "Get a backup with date and time first. Select folder inventory, appropriate stream for data collection and template transfer from files.")
 Case "guideButton3|label"
 HAMU_UITextPart7 = HAMU_L("Rehberi aç", "Open Guide")
 Case "guide4|label"
 HAMU_UITextPart7 = HAMU_L("05  Görseller ve kodlar", "05  Images and Codes")
 Case "guide4|helperText"
 HAMU_UITextPart7 = HAMU_L("Hücre içine veya üstüne resim ekleme seçeneklerini öğrenin. Barkod ve karekod üretiminde internet ve veri paylaşımı bilgilerini okuyun.", "Learn the options of adding pictures into the cell or above. Read internet and data sharing information in barcode and square code production.")
 Case "guideButton4|label"
 HAMU_UITextPart7 = HAMU_L("Rehberi aç", "Open Guide")
 Case "brand|label"
 HAMU_UITextPart7 = HAMU_L("HAMU · V1", "HAMU · V1")
 Case "brand|helperText"
 HAMU_UITextPart7 = HAMU_L("Excel üretkenlik ve veri araçları", "Excel productivity and data tools")
 Case "brandLogo|altText"
 HAMU_UITextPart7 = HAMU_L("Özgün HAMU logosu", "Original HAMU logo")
 Case "brandName|label"
 HAMU_UITextPart7 = HAMU_L("Hüseyin Avni UZUN", "Hüseyin Avni UZUN")
 Case "brandPurpose|label"
 HAMU_UITextPart7 = HAMU_L("Veriyi düzenleyin. Tekrarlayan işleri hızlandırın. Sonuçları anlaşılır kılın.", "Organize data. Speed up repetitive tasks. Make results clear.")
 Case "profile0|label"
 HAMU_UITextPart7 = HAMU_L("Web sitesi", "Website")
 Case "profile1|label"
 HAMU_UITextPart7 = HAMU_L("GitHub", "GitHub")
 Case "profile2|label"
 HAMU_UITextPart7 = HAMU_L("hamu@huseyinavniuzun.com", "hamu@huseyinavniuzun.com")
 Case "settingsCenter|label"
 HAMU_UITextPart7 = HAMU_L("Ayarlar", "Settings")
 Case "settingsCenter|helperText"
 HAMU_UITextPart7 = HAMU_L("Dil, kayıt konumları ve işlem tercihlerini bu merkezden yönetin.", "Manage language, registration locations and trading preferences from this center.")
 Case "openSettings|label"
 HAMU_UITextPart7 = HAMU_L("Tüm ayarları düzenle", "Edit All Settings")
 Case "settingLanguage|label"
 HAMU_UITextPart7 = HAMU_L("Arayüz dili", "Interface Language")
 Case "settingsStorage|label"
 HAMU_UITextPart7 = HAMU_L("Tercihler kullanıcıya özeldir; diğer çalışma kitaplarının ayarlarını değiştirmez.", "Preferences belong to this Windows user and do not change other workbooks' settings.")
 Case "learn|label"
 HAMU_UITextPart7 = HAMU_L("İhtiyacınız olan bilgi", "Find the Information You Need")
 Case "topic1|label"
 HAMU_UITextPart7 = HAMU_L("Aralık ve çıktı seçimi", "Range and Output Selection")
 Case "topic6|label"
 HAMU_UITextPart7 = HAMU_L("HAMU hücre fonksiyonları", "HAMU Worksheet Functions")
 Case "topic7|label"
 HAMU_UITextPart7 = HAMU_L("Sorun giderme ve sınırlar", "Troubleshooting and Limits")
 Case "topic8|label"
 HAMU_UITextPart7 = HAMU_L("Tüm araçlar", "All Tools")
 Case "aboutProfile|label"
 HAMU_UITextPart7 = HAMU_L("HAMU hakkında", "About HAMU")
 Case "note|label"
 HAMU_UITextPart7 = HAMU_L("Çalışmaya başlamadan", "Before You Start")
 Case "backupNote|label"
 HAMU_UITextPart7 = HAMU_L("VBA işlemleri geri alma geçmişini temizleyebilir. Önemli dosyalarda önce Yedek Al kullanın.", "VBA operations can clear Excel's undo history. Back up important files before proceeding.")
 Case "internetNote|label"
 HAMU_UITextPart7 = HAMU_L("Barkod ve karekod içeriği, üretim için üçüncü taraf internet servisine gönderilir.", "Barcode and QR content is sent to a third-party internet service for generation.")
 Case "textSub0|label"
 HAMU_UITextPart7 = HAMU_L("Metin Araçları", "Text Tools")
 Case "textSub1|label"
 HAMU_UITextPart7 = HAMU_L("Normal İfade", "Regular Expressions")
 Case "textSub2|label"
 HAMU_UITextPart7 = HAMU_L("Hücre Düzeni", "Cell Layout")
 Case "r_HAMU_DataEntry|label"
 HAMU_UITextPart7 = HAMU_L("Veri Girişi", "Data Entry")
 Case "r_HAMU_DataEntry|screentip"
 HAMU_UITextPart7 = HAMU_L("Veri Girişi", "Data Entry")
 Case "r_HAMU_DataEntry|supertip"
 HAMU_UITextPart7 = HAMU_L("Excel’in yerleşik veri giriş formunu açar. Bir tabloya veya benzersiz başlıkları olan kesintisiz listeye tıklayın; en fazla 32 sütun.", "Open Excel's native data form for the current table or contiguous list with unique headers; at most 32 columns.")
 End Select
End Function

Private Function HAMU_UITextPart8(ByVal key As String) As String
 Select Case key
 Case "r_HAMU_ClearHighlights|label"
 HAMU_UITextPart8 = HAMU_L("Vurguyu Kaldır", "Clear Highlights")
 Case "r_HAMU_ClearHighlights|screentip"
 HAMU_UITextPart8 = HAMU_L("Vurguyu Kaldır", "Clear Highlights")
 Case "r_HAMU_ClearHighlights|supertip"
 HAMU_UITextPart8 = HAMU_L("Seçili alandaki HAMU vurgu renklerini kaldırır. Aynı oturumda önceki dolguyu geri yükler; daha eski vurgular sarı/turuncu paletiyle tanınır.", "Remove HAMU highlights in the selection. Restore prior fills in the same session; older highlights are recognized by their yellow/orange palette.")
 Case "auditHighlight|label"
 HAMU_UITextPart8 = HAMU_L("Vurgu ve Sayım", "Highlight and Count")
 Case "worksheetFunctions|label"
 HAMU_UITextPart8 = HAMU_L("Hücre Fonksiyonları", "Worksheet Functions")
 Case "lambdaFunctions|label"
 HAMU_UITextPart8 = HAMU_L("LAMBDA Şablonları", "LAMBDA Templates")
 Case "openLambdaTemplates|label"
 HAMU_UITextPart8 = HAMU_L("Şablonu Aç", "Open Template")
 Case "openLambdaTemplates|supertip"
 HAMU_UITextPart8 = HAMU_L("Üç uzman LAMBDA tanımı içeren dosyayı açar. Aktarma Sihirbazı ile kendi kitabınıza alın.", "Open the template containing three expert LAMBDA functions. Use Transfer Wizard to copy them into your workbook.")
 Case "r_UDF_LambdaMultiLookup|label"
 HAMU_UITextPart8 = HAMU_L("HAMU_L_CokluAra", "HAMU_L_CokluAra")
 Case "r_UDF_LambdaMultiLookup|screentip"
 HAMU_UITextPart8 = HAMU_L("HAMU_L_CokluAra", "HAMU_L_CokluAra")
 Case "r_UDF_LambdaMultiLookup|supertip"
 HAMU_UITextPart8 = HAMU_L("İki ayrı anahtar/sonuç sütununda tam eşleşme arar; ilk sonucu verir. Örnek: =HAMU_L_CokluAra(G2;A2:A10;B2:B10;D2:D10;E2:E10). LAMBDA şablonundan aktarılır; Microsoft 365/Excel 2024 gerekir.", "Exact lookup across two key/result pairs, returning the first match. Example: =HAMU_L_CokluAra(G2;A2:A10;B2:B10;D2:D10;E2:E10). Transfer it from the LAMBDA template; Microsoft 365/Excel 2024 required.")
 Case "r_UDF_LambdaMultiSum|label"
 HAMU_UITextPart8 = HAMU_L("HAMU_L_CokluTopla", "HAMU_L_CokluTopla")
 Case "r_UDF_LambdaMultiSum|screentip"
 HAMU_UITextPart8 = HAMU_L("HAMU_L_CokluTopla", "HAMU_L_CokluTopla")
 Case "r_UDF_LambdaMultiSum|supertip"
 HAMU_UITextPart8 = HAMU_L("İki ayrı anahtar/değer sütununda koşula uyan sayıları toplar. Örnek: =HAMU_L_CokluTopla(G2;A2:A10;B2:B10;D2:D10;E2:E10). Şablondan aktarılır; Microsoft 365/Excel 2024 gerekir.", "Sum matching values across two key/value pairs. Example: =HAMU_L_CokluTopla(G2;A2:A10;B2:B10;D2:D10;E2:E10). Transfer from the template; Microsoft 365/Excel 2024 required.")
 Case "r_UDF_LambdaUniqueMerge|label"
 HAMU_UITextPart8 = HAMU_L("HAMU_L_BenzersizBirlestir", "HAMU_L_BenzersizBirlestir")
 Case "r_UDF_LambdaUniqueMerge|screentip"
 HAMU_UITextPart8 = HAMU_L("HAMU_L_BenzersizBirlestir", "HAMU_L_BenzersizBirlestir")
 Case "r_UDF_LambdaUniqueMerge|supertip"
 HAMU_UITextPart8 = HAMU_L("İki alandaki dolu değerleri tek bir sıralı benzersiz listeye dönüştürür. Örnek: =HAMU_L_BenzersizBirlestir(A2:B10;D2:E10). Sonuç için boş taşma alanı gerekir. Microsoft 365/Excel 2024 gerekir.", "Combine nonblank values from two areas into a sorted unique list. Example: =HAMU_L_BenzersizBirlestir(A2:B10;D2:E10). Requires a blank spill area and Microsoft 365/Excel 2024.")
 Case "fileHub|label"
 HAMU_UITextPart8 = HAMU_L("Dosya", "Files")
 Case "r_HAMU_Backup|label"
 HAMU_UITextPart8 = HAMU_L("Yedek Oluştur", "Create Backup")
 Case "r_HAMU_Backup|screentip"
 HAMU_UITextPart8 = HAMU_L("Yedek Oluştur", "Create Backup")
 Case "r_HAMU_Backup|supertip"
 HAMU_UITextPart8 = HAMU_L("Dosyanın tarih ve saatli kopyasını Belgeler > HAMU > Backup klasörüne kaydeder. Sonuç ekranında tam yolu gösterir; Klasörü Aç ile yedek klasörüne ulaşabilirsiniz.", "Save a timestamped copy to the backup folder configured in Settings. The result displays its full path and an Open Folder button.")
 Case "r_HAMU_SaveCopy|label"
 HAMU_UITextPart8 = HAMU_L("Kopya Kaydet", "Save Copy")
 Case "r_HAMU_SaveCopy|screentip"
 HAMU_UITextPart8 = HAMU_L("Kopya Kaydet", "Save Copy")
 Case "r_HAMU_SaveCopy|supertip"
 HAMU_UITextPart8 = HAMU_L("Dosyanın bir kopyasını seçtiğiniz konuma kaydeder. Açık dosyanın adı, konumu ve kaydedilmemiş değişiklik durumu korunur; biçim ve makrolar değiştirilmez.", "Save a copy at a chosen location, preserving the open workbook identity, unsaved state, file format and macros.")
 Case "r_HAMU_ExportSheetsAsPDF|label"
 HAMU_UITextPart8 = HAMU_L("PDF Kaydet", "Save PDF")
 Case "r_HAMU_ExportSheetsAsPDF|screentip"
 HAMU_UITextPart8 = HAMU_L("PDF Kaydet", "Save PDF")
 Case "r_HAMU_ExportSheetsAsPDF|supertip"
 HAMU_UITextPart8 = HAMU_L("Alan Seç veya Etkin Sayfa seçeneğiyle PDF kaydeder. Alan Seç her zaman Excel aralık seçicisini açar; mevcut seçimi kendiliğinden kaydetmez. Otomatik/dikey/yatay yön sunar.", "Save PDF using Select Range or Active Sheet. Select Range always opens Excel's range picker instead of automatically exporting the selection. Choose automatic, portrait or landscape.")
 Case "mergeHub|label"
 HAMU_UITextPart8 = HAMU_L("Birleştir", "Combine")
 Case "r_HAMU_SmartAppendByHeaders|label"
 HAMU_UITextPart8 = HAMU_L("Akıllı Birleştir", "Smart Append")
 Case "r_HAMU_SmartAppendByHeaders|screentip"
 HAMU_UITextPart8 = HAMU_L("Başlıklara Göre Akıllı Birleştir", "Smart Append")
 Case "r_HAMU_SmartAppendByHeaders|supertip"
 HAMU_UITextPart8 = HAMU_L("Aynı bilgileri içeren sayfaları alt alta birleştirir. Örneğin Ocak ve Şubat satışlarında sütun sırası farklı olsa da Ürün başlıkları eşleştirilir. Eksik sütunlar boş kalır, KaynakSayfa eklenir. Kaynak sayfalar değiştirilmez; ilk satırlar başlık olmalıdır.", "Append worksheets by matching header names, even when column order differs. Missing columns remain blank. Add the source sheet name and preserve the source worksheets.")
 Case "extractHub|label"
 HAMU_UITextPart8 = HAMU_L("Ayıkla", "Extract")
 Case "r_HAMU_SplitDataWizard|label"
 HAMU_UITextPart8 = HAMU_L("Veriyi Böl", "Split Data")
 Case "r_HAMU_SplitDataWizard|screentip"
 HAMU_UITextPart8 = HAMU_L("Veriyi Böl", "Split Data")
 Case "r_HAMU_SplitDataWizard|supertip"
 HAMU_UITextPart8 = HAMU_L("Seçili tabloyu satır sayısına, eşit parçalara, tek veya çoklu sütun değerlerine göre ya da yıl, çeyrek, ay ve hafta bazında tarih grupları oluşturarak böler. Çıktıları yeni çalışma sayfalarına veya ayrı Excel dosyalarına aktarabilirsiniz.", "Split values or datasets using the options shown in the dialog. Choose the source columns and desired output.")
 Case "r_HAMU_ExtractUniqueList|label"
 HAMU_UITextPart8 = HAMU_L("Benzersiz Liste", "Unique List")
 Case "r_HAMU_ExtractUniqueList|screentip"
 HAMU_UITextPart8 = HAMU_L("Benzersiz Liste", "Unique List")
 Case "r_HAMU_ExtractUniqueList|supertip"
 HAMU_UITextPart8 = HAMU_L("Tekrar eden değerleri bir kez listeler. Çıktı için mevcut sayfada veya başka sayfada başlangıç hücresi seçebilir ya da yeni sayfa oluşturabilirsiniz. Dolu hedefte onay ister; kaynak hata ve boş hücrelerini atlar.", "Create a list of distinct values in a selected destination range or on a new worksheet. Configure the default destination behavior in Settings.")
 Case "reshapeHub|label"
 HAMU_UITextPart8 = HAMU_L("Dönüştür", "Reshape")
 Case "r_HAMU_GroupAndSummarize|label"
 HAMU_UITextPart8 = HAMU_L("Grupla ve Özetle", "Group and Summarize")
 Case "r_HAMU_GroupAndSummarize|screentip"
 HAMU_UITextPart8 = HAMU_L("Grupla ve Özetle", "Group and Summarize")
 Case "r_HAMU_GroupAndSummarize|supertip"
 HAMU_UITextPart8 = HAMU_L("Bir veya birden fazla sütuna göre kayıtları gruplar; birden fazla sayısal sütunu aynı anda toplar. Sütunları numarayla veya başlık adıyla seçebilirsiniz.", "Group records by one or more key columns and summarize selected numeric fields. Specify column numbers or header names.")
 Case "qualityHub|label"
 HAMU_UITextPart8 = HAMU_L("Kalite", "Quality")
 Case "extraFunctions|label"
 HAMU_UITextPart8 = HAMU_L("Fonksiyonlar", "Functions")
 Case "textSub1|screentip"
 HAMU_UITextPart8 = HAMU_L("Normal İfade", "Regular Expressions")
 End Select
End Function

Private Function HAMU_UITextPart9(ByVal key As String) As String
 Select Case key
 Case "textSub1|supertip"
 HAMU_UITextPart9 = HAMU_L("Normal ifadelerle metin bulun, ayıklayın veya değiştirin. Hazır örnekler ve özel desenler kullanılabilir.", "Find, extract or replace text using regular expressions. Use examples or a custom pattern.")
 Case "textSub2|screentip"
 HAMU_UITextPart9 = HAMU_L("Hücre Düzeni", "Cell Layout")
 Case "textSub2|supertip"
 HAMU_UITextPart9 = HAMU_L("Hücre görünümünü, sayı ve tarih biçimlerini ve sütun genişliklerini düzenleyen araçlar.", "Tools for cell layout, number and date formats, and column widths.")
 Case "extraFunctions|screentip"
 HAMU_UITextPart9 = HAMU_L("Fonksiyonlar", "Functions")
 Case "extraFunctions|supertip"
 HAMU_UITextPart9 = HAMU_L("HAMU hücre fonksiyonlarını ve LAMBDA şablonlarını örnekleriyle inceleyin.", "Explore HAMU worksheet functions and LAMBDA templates with examples.")
 Case "r_HAMU_ExpandColumns|label"
 HAMU_UITextPart9 = HAMU_L("Genişlet", "Fit Width")
 Case "r_HAMU_ExpandColumns|screentip"
 HAMU_UITextPart9 = HAMU_L("Genişlet", "Fit Width")
 Case "r_HAMU_ExpandColumns|supertip"
 HAMU_UITextPart9 = HAMU_L("Seçili alanın sütun genişliğini metne göre ayarlar.", "Fit selected column widths to the content.")
 Case "r_HAMU_TableTotals|label"
 HAMU_UITextPart9 = HAMU_L("Toplam Satırı", "Total Row")
 Case "r_HAMU_TableTotals|screentip"
 HAMU_UITextPart9 = HAMU_L("Toplam Satırı", "Total Row")
 Case "r_HAMU_TableTotals|supertip"
 HAMU_UITextPart9 = HAMU_L("Etkin tablonun toplam satırını açar veya kapatır; sayısal sütunlara toplam ekler.", "Toggle the active table total row and sum numeric columns.")
 Case "r_HAMU_CopyVisibleTable|label"
 HAMU_UITextPart9 = HAMU_L("Görünür Satırları Kopyala", "Copy Visible Rows")
 Case "r_HAMU_CopyVisibleTable|screentip"
 HAMU_UITextPart9 = HAMU_L("Görünür Satırları Kopyala", "Copy Visible Rows")
 Case "r_HAMU_CopyVisibleTable|supertip"
 HAMU_UITextPart9 = HAMU_L("Filtre sonrası görünen kayıtları yeni sayfada bağımsız tabloya kopyalar.", "Copy visible filtered records into a standalone table on a new sheet.")
 End Select
End Function
