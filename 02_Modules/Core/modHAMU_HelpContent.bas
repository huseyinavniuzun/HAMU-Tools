Attribute VB_Name = "modHAMU_HelpContent"
Option Explicit
Option Private Module
Public Function HAMU_GuideTitle(ByVal topic As String) As String
 Select Case topic
 Case "start": HAMU_GuideTitle = HAMU_Text("İlk adımlar")
 Case "selection": HAMU_GuideTitle = HAMU_Text("Aralık ve çıktı seçimi")
 Case "clean": HAMU_GuideTitle = HAMU_Text("Metin ve veri temizleme")
 Case "tables": HAMU_GuideTitle = HAMU_Text("Tablolar ve özetleme")
 Case "files": HAMU_GuideTitle = HAMU_Text("Dosyalar ve yedekleme")
 Case "visual": HAMU_GuideTitle = HAMU_Text("Resimler, barkod ve karekod")
 Case "functions": HAMU_GuideTitle = HAMU_Text("HAMU hücre fonksiyonları")
 Case "errors": HAMU_GuideTitle = HAMU_Text("Sorun giderme ve sınırlar")
 Case "catalog": HAMU_GuideTitle = HAMU_Text("Tüm araçlar")
 End Select
End Function
Public Function HAMU_GuideBody(ByVal topic As String) As String
 Select Case topic
 Case "start"
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("HAMU Tools, Excel içinde dosya, veri ve hücre işlemlerini hızlandırır.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("1. İşlem yapacağınız kitabı ve sayfayı açın." & vbLf & "2. Kaynak veriyi seçin; başlık isteyen araçlarda başlık satırını da ekleyin." & vbLf & "3. Şeritten uygun aracı açın, açıklamasını okuyun ve seçenekleri belirleyin." & vbLf & "4. Yeni çıktı isteyen araçlarda boş bir hedef seçin; sonuç ve satır sayısını kontrol edin.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Önerilen akış: Yedek Al " & ChrW(8594) & " Başlıkları Normalize Et " & ChrW(8594) & " Veri Kalitesi " & ChrW(8594) & " Tablo Oluştur " & ChrW(8594) & " Grupla veya Unpivot.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Komutların tamamı şerittedir. Dosya > HAMU Tools ise iş akışlarını ve bu rehberi sunar. Çalışma kitabı açık değilken yardım ve hakkında ekranlarına erişebilirsiniz.") & vbCrLf & vbCrLf
 Case "selection"
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Mevcut seçiminiz, aralık kutusunda başlangıç adresi olarak görünür. Kutudan farklı hücreleri seçebilir veya geçerli bir Excel adresi yazabilirsiniz. Seçimi Enter ile onaylayın; X ile kapatmak işlemi iptal eder.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Kaynak, okunacak veridir; hedef ise sonucun yazılacağı yerdir. Birden fazla giriş isteyen araçlarda kutu başlığını kontrol edin. Tek hedef hücresi, liste sonuçlarında başlangıç noktasıdır; sonuç aşağıya veya sağa genişleyebilir.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Başlıklar: Tablo, birleştirme, grupla ve pivot araçlarında ilk satırın anlamlı ve benzersiz başlıklar taşıması gerekir.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Tarih Seç, önceden seçili aktif hücreye yazar. Formdaki ay ve yılı değiştirebilir, Bugün ile güncel tarihe dönebilirsiniz. Seçilen tarihin bugüne uzaklığı altta gösterilir.") & vbCrLf & vbCrLf
 Case "clean"
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Önce küçük bir örnekte deneyin, ardından tüm listeye uygulayın.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Başlıkları Normalize Et: Boşluk ve yinelenen başlıkları düzenler. Metin dönüşümleri: Büyük/küçük harf, boşluk, ön/son ek ve metin parçası araçlarını kullanın. Normal İfade: Şablon seçin veya özel desen yazın; açıklama ve örnekleri inceleyin.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Boşlukları Sil, boş hücreleri kaldırıp alttaki hücreleri yukarı taşır; satırlar arasında ilişkili sütunlar varsa sonuçları dikkatle kontrol edin. Boş hücreleri sıfırla doldurmak hücre silmez.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Standart Sayı, tam sayı görüntüsü uygular; Standart Tarih, tarih görünümünü düzenler. Hücrenin gerçek türü ile görüntü biçimi farklıdır. Baştaki Sıfırlar, kodları metin olarak saklar; hesaplama yapılacak sayılarda kullanmayın.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Veri Kalitesi araçlarıyla boşlukları, hataları ve tekrarları inceleyin; denetim sonucu bir temizleme kararına yardımcı olur.") & vbCrLf & vbCrLf
 Case "tables"
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Tablo Oluştur: Başlıkları içeren veri alanından Excel tablosu oluşturur. Otomatik Sığdır: Satır ve sütun boyutlarını içeriğe göre düzenler. Tabloyu Genişlet: İmleç bir Excel tablosunun içindeyken kullanılır.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Sütun Seçerek Yeni Tablo: İhtiyacınız olan sütunlardan ayrı çıktı üretir. Grupla: Bir veya daha fazla anahtar sütuna göre özet oluşturur. Pivot: Kategorileri sütunlara taşır. Unpivot: Geniş sütun yapısını alan/değer satırlarına dönüştürür; 2026.01 gibi metin başlıkları korunur.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Akıllı Birleştir aynı yapılı listeleri alt alta toplar. Klasörlerden Veri Topla birden fazla dosyayı okur; dosya ve sayfa yapısını önceden kontrol edin. Anahtarla Eşleştir, ortak anahtara göre ayrı listeleri ilişkilendirir. Anahtarların türü, boşlukları ve tekrarları eşleşmeyi etkiler.") & vbCrLf & vbCrLf
 Case "files"
  HAMU_GuideBody = HAMU_FilesGuide()
 Case "visual"
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Barkod/Karekod ekranında hedef hücreyi ve içerik modelini seçin. Metin, bağlantı veya e-posta gibi seçenekler aynı içeriğin doğru kod biçiminde hazırlanmasını sağlar. E-posta seçimi ileti göndermez; kodu tarayan kişinin uygulaması bir taslak açabilir.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Kod üretimi internet ve üçüncü taraf servis kullanır; kod içeriği üretim servisine gönderilir. Gizli veya kişisel veri girmeden önce bunu değerlendirin. Bağlantı veya servis sorunu varsa kod üretilemeyebilir.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Klasörden Resim Ekle: Hücre içine veya hücre üstüne yerleştirme seçeneğini belirleyin. Hücre içine özellik desteği Excel sürümüne bağlıdır; hücre üstündeki resimler hücre sınırlarına göre yerleştirilir.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Satırları Renklendir ve sayfa/görsel araçları ilgili gruptadır. Çalışma sayfası renk formülleri doğrudan dolgu rengini okur; denetim komutları görüntülenen rengi, koşullu biçimlendirme dahil, kullanır.") & vbCrLf & vbCrLf
 Case "functions"
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Formülleri hücreye =HAMU_ yazarak bulun. Aşağıdaki örneklerde Türkçe Excel ayırıcıları kullanılmıştır; Excel diliniz farklıysa ayırıcıyı uyarlayın.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("=HAMU_URLCoz(A1) | UTF-8 URL karakterlerini çözer. Geçersiz yüzde dizisini korur.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("=HAMU_RengeGoreTopla(A1:A20;B1) | B1 ile aynı dolgu rengine sahip sayıları toplar. Biçim değişimi sonrası F9 kullanın.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("=HAMU_RengeGoreSay(A1:A20;B1) | B1 ile aynı dolgu rengine sahip hücreleri sayar. Biçim değişimi sonrası F9 kullanın.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("=HAMU_YerelDosyaYolu(A1) | Kişisel OneDrive URL yolunu yerel OneDrive ortam değişkeniyle eşler; diğer URL türlerini korur.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("=HAMU_GizliLink(A1) | Hücrenin ilk köprü adresini getirir.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("=HAMU_LinkParamDegeri(A1;""id"") | URL içindeki sorgu parametresinin değerini getirir.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("=HAMU_SayidanMetine(123) | Türkçe: yüz yirmi üç. 15 basamaktan küçük mutlak değerler; ondalık kısmı tek tek okur.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("=HAMU_TablodanVeriGetir(A1;""Tablo1"";""Kod"";""Ad"") | Formülün bulunduğu kitapta tablodan ilk eşleşen değeri getirir.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Renk değiştirdikten sonra F9 ile yeniden hesaplayın. Renk formülleri otomatik bir renk değişimi olayı izlemez. HAMU eklentisi kapalıysa bu özel fonksiyonlar hesaplanamaz; dosyayı paylaşacağınız kişinin de eklentiye ihtiyacı olabilir.") & vbCrLf & vbCrLf
 Case "errors"
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Komut pasif veya çalışmıyor: Açık bir çalışma kitabı ve uygun sayfa olduğundan emin olun. Hücre düzenleme modundan Enter/Esc ile çıkın.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Aralık geçersiz: Doğru kitap/sayfayı seçin; aracın tek alan, tek sütun veya başlık şartını kontrol edin. Korunan sayfalarda yazma işlemleri engellenebilir.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Tablo bulunamadı: İmleci Excel tablosunun içine taşıyın veya önce Tablo Oluştur çalıştırın.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Formül sonucu beklenmiyor: Sayı/metin türünü, başlık adlarını, anahtar tekrarlarını ve F9 hesaplamasını kontrol edin.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Kod/resim oluşturulamadı: İnternet, servis erişimi ve Excel sürüm desteğini kontrol edin.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("Uzun klasör ve veri işlemlerinde Excel bekleyebilir; alt klasör kapsamını daraltın ve küçük örnekle başlayın. Sorun bildirirken araç adı, Excel sürümü, hata metni ve kişisel veri içermeyen bir örnek paylaşın.") & vbCrLf & vbCrLf
  HAMU_GuideBody = HAMU_GuideBody & HAMU_Text("HAMU Tools V1, masaüstü Excel VBA eklentisidir; web Excel VBA çalıştırmaz.") & vbCrLf & vbCrLf
 Case "catalog"
  HAMU_GuideBody = HAMU_CatalogGuide()
 End Select
 If topic = "visual" Then HAMU_GuideBody = HAMU_GuideBody & vbCrLf & HAMU_L("Kod içeriğinin altındaki Toplu giriş alanından tek sütun seçebilirsiniz. Kodlar verilerin sağındaki ilk boş sütuna, kaynak satırlarıyla aynı hizaya gelir. Başlık seçmeyin; en fazla 250 kod. Resim Kaydet PNG/JPG/GIF sunar.", "Use Batch input below Code content to select one column. Codes go in the first empty column to the right, aligned with source rows. Exclude headers; at most 250 codes. Save Ima" & _
        "ge offers PNG/JPG/GIF.")
 If topic = "files" Then HAMU_GuideBody = HAMU_GuideBody & vbCrLf & HAMU_L("PDF Kaydet seçili aralık veya etkin sayfa için çalışır; otomatik/dikey/yatay yön seçilir. Tarihli Sürüm eski dosyayı doğrulanmış yeni kopyasıyla değiştirir.", "Save PDF exports the selection or active sheet with automatic, portrait or landscape orientation. Dated Version replaces the original only after the new copy has been verified.")
 If topic = "functions" Then HAMU_GuideBody = HAMU_GuideBody & vbCrLf & HAMU_L("Üç yeni gerçek LAMBDA: HAMU_L_CokluAra, HAMU_L_CokluTopla ve HAMU_L_BenzersizBirlestir. Bilgi > LAMBDA Şablonları > Şablonu Aç ile örneklere ulaşın; Aktarma Sihirbazı ile kitabınıza aktarın. Bunlar VBA fonksiyonlarından ayrıdır ve Microsoft 365/Excel 2024 gerektirir.", "Three native LAMBDA functions: HAMU_L_CokluAra, HAMU_L_CokluTopla and HAMU_L_BenzersizBirlestir. Open Information > LAMBDA Templates > Open Template for example" & _
        "s, then use Transfer Wizard. They are separate from the VBA worksheet functions and require Microsoft 365/Excel 2024.")
 If topic = "tables" Then HAMU_GuideBody = HAMU_GuideBody & vbCrLf & HAMU_L("Veri Girişi Excel’in kendi formunu açar; en fazla 32 sütun ve benzersiz başlık gerekir. Genişlet sütun genişliğini, Sığdır satır yüksekliğini ayarlar. Tabloyu Genişlet bitişik veri satırlarını tabloya dahil eder. Tablo Araçları içindeki Toplam Satırı toplamları açar veya kapatır; Görünür Satırları Kopyala filtrelenen kayıtları yeni tabloya alır. Sütun Seç, başlık adlarını veya seçim içindeki 1,2,3 numaralarını kabul eder.", "Data Entry opens Excel's native form; at most 32 columns and unique headers are required. AutoFit adjusts layout; Extend Table includes data rows. Select Column" & _
        "s accepts header names or positions 1,2,3 within the selection.")

 If topic = "files" Then HAMU_GuideBody = HAMU_GuideBody & vbCrLf & HAMU_L("PDF: Alan Seç ile adresi açıkça seçin; Etkin Sayfa tüm veri alanını alır. Kopya Kaydet seçilen konuma kopya oluşturur, açık dosyayı değiştirmez.", "PDF: Select Range explicitly picks the address; Active Sheet exports the data area. Save Copy writes a separate copy without changing the open workbook.")
End Function
Public Sub HAMU_OpenGuide(Optional ByVal topic As String = "start")
 Dim view As frmHAMU_Help
 Set view = New frmHAMU_Help
 view.Setup topic
 view.show vbModal
 Unload view
End Sub

Public Function HAMU_SearchGuide(ByVal query As String) As String
 Dim topic As Variant, id As Variant, body As String, title As String, key As String, text As String, match As Long, result As String
 key = HAMU_Data_ToSnakeCase(query)
 If Len(key) = 0 Then HAMU_SearchGuide = HAMU_L("Aranacak kelimeyi yazın.", "Enter a search term."): Exit Function
 For Each topic In Split("start|selection|clean|tables|files|visual|functions|errors", "|")
  title = HAMU_GuideTitle(CStr(topic)): body = HAMU_GuideBody(CStr(topic))
  text = HAMU_Data_ToSnakeCase(title & " " & body)
  If InStr(1, text, key, vbTextCompare) > 0 Then result = result & title & vbCrLf & Left$(body, 450) & vbCrLf & vbCrLf
 Next
 For Each id In Split("HAMU_Transfer|HAMU_Backup|HAMU_AddDateToFilename|HAMU_ExportSheetsAsPDF|HAMU_FolderIndex|HAMU_CombineCSVs|HAMU_ConsolidateSheets|HAMU_AppendSelectedTables|HAMU_" & _
        "SmartAppendByHeaders|HAMU_CollectFilesData|HAMU_SmartJoin|HAMU_GroupAndSummarize|HAMU_UnpivotColumnsToRows|HAMU_SplitDataWizard|HAMU_SplitByValueToFiles|HAMU_Sp" & _
        "litURLParameters|HAMU_PivotExport|HAMU_PivotRowsToColumns|HAMU_FillDownBlanks|HAMU_MergeDuplicates|HAMU_StandardizeHeaders|HAMU_ConvertColumnTypes|HAMU_SelectCo" & _
        "lumnsToNewTable|HAMU_FilterRowsToNewTable|HAMU_SampleData|HAMU_ShuffleRows|HAMU_CompareTwoLists|HAMU_ExtractUniqueList|HAMU_DataProfile|HAMU_DataQualityReport|H" & _
        "AMU_FindMissingCombinations|HAMU_CleanData|HAMU_TextCase|HAMU_PatternFind|HAMU_RegexReplace|HAMU_ExtractDigits|HAMU_ConvertQuotedNumbers|HAMU_MaskData|HAMU_Conv" & _
        "ertFormulasToValues|HAMU_UnmergeAndFill|HAMU_StandardFormat|HAMU_DeleteBlankCells|HAMU_DeleteDropdowns|HAMU_RowStriping|HAMU_DatePicker|HAMU_FixDates|HAMU_Check" & _
        "Dates|HAMU_DateBuilder|HAMU_SheetIndex|HAMU_SortSheets|HAMU_ShowHidden|HAMU_DynamicValidation|HAMU_SelectionToPNG|HAMU_InsertPhotosFromFolder|HAMU_CreateBarcode" & _
        "|HAMU_CreateQRCode|HAMU_CompareSheets|HAMU_MarkProblemCells|HAMU_HighlightByCriteria|HAMU_HighlightDuplicates|HAMU_AddDuplicateCount|HAMU_ListLinks|HAMU_BreakLi" & _
        "nks|HAMU_Help|HAMU_About|HAMU_AutoFit|HAMU_AutoTable|HAMU_ResizeTable|HAMU_UDFHelp|UDF_URLDecode|UDF_RenklileriTopla|UDF_RengeGoreSay|UDF_YerelDosyaYolu|UDF_Giz" & _
        "liLink|UDF_LinkParamDegeri|UDF_SayidanMetine|UDF_TablodanVeriGetir|HAMU_CountByColor|HAMU_SumByColor|HAMU_FillBlanksZero|HAMU_AddTextAffixes|HAMU_LeadingZeros|H" & _
        "AMU_DeduplicateItems|HAMU_ExtractTextPart|HAMU_DateRangeList|HAMU_FormulaAudit|HAMU_SelectSpecialCells|HAMU_ListDefinedNames|HAMU_DataEntry|HAMU_ClearHighlights" & _
        "|UDF_LambdaMultiLookup|UDF_LambdaMultiSum|UDF_LambdaUniqueMerge|HAMU_SaveCopy", "|")
  body = HAMU_CommandIntro(CStr(id))
  If InStr(1, HAMU_Data_ToSnakeCase(body & " " & HAMU_CommandLabel(CStr(id))), key, vbTextCompare) > 0 Then result = result & HAMU_CommandLabel(CStr(id)) & vbCrLf & body & vbCrLf & vbCrLf
 Next
 If Len(result) = 0 Then result = HAMU_L("Sonuç bulunamadı. Daha kısa bir kelime deneyin.", "No results. Try a shorter search term.")
 HAMU_SearchGuide = result
End Function

Private Function HAMU_CatalogGuide() As String
 HAMU_CatalogGuide = "HAMU Tools V1" & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_Transfer") & vbCrLf & HAMU_CommandIntro("HAMU_Transfer") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_Backup") & vbCrLf & HAMU_CommandIntro("HAMU_Backup") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_AddDateToFilename") & vbCrLf & HAMU_CommandIntro("HAMU_AddDateToFilename") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_ExportSheetsAsPDF") & vbCrLf & HAMU_CommandIntro("HAMU_ExportSheetsAsPDF") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_FolderIndex") & vbCrLf & HAMU_CommandIntro("HAMU_FolderIndex") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_CombineCSVs") & vbCrLf & HAMU_CommandIntro("HAMU_CombineCSVs") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_ConsolidateSheets") & vbCrLf & HAMU_CommandIntro("HAMU_ConsolidateSheets") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_AppendSelectedTables") & vbCrLf & HAMU_CommandIntro("HAMU_AppendSelectedTables") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_SmartAppendByHeaders") & vbCrLf & HAMU_CommandIntro("HAMU_SmartAppendByHeaders") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_CollectFilesData") & vbCrLf & HAMU_CommandIntro("HAMU_CollectFilesData") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_SmartJoin") & vbCrLf & HAMU_CommandIntro("HAMU_SmartJoin") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_GroupAndSummarize") & vbCrLf & HAMU_CommandIntro("HAMU_GroupAndSummarize") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_UnpivotColumnsToRows") & vbCrLf & HAMU_CommandIntro("HAMU_UnpivotColumnsToRows") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_SplitDataWizard") & vbCrLf & HAMU_CommandIntro("HAMU_SplitDataWizard") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_SplitByValueToFiles") & vbCrLf & HAMU_CommandIntro("HAMU_SplitByValueToFiles") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_SplitURLParameters") & vbCrLf & HAMU_CommandIntro("HAMU_SplitURLParameters") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_PivotExport") & vbCrLf & HAMU_CommandIntro("HAMU_PivotExport") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_PivotRowsToColumns") & vbCrLf & HAMU_CommandIntro("HAMU_PivotRowsToColumns") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_FillDownBlanks") & vbCrLf & HAMU_CommandIntro("HAMU_FillDownBlanks") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_MergeDuplicates") & vbCrLf & HAMU_CommandIntro("HAMU_MergeDuplicates") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_StandardizeHeaders") & vbCrLf & HAMU_CommandIntro("HAMU_StandardizeHeaders") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_ConvertColumnTypes") & vbCrLf & HAMU_CommandIntro("HAMU_ConvertColumnTypes") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_SelectColumnsToNewTable") & vbCrLf & HAMU_CommandIntro("HAMU_SelectColumnsToNewTable") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_FilterRowsToNewTable") & vbCrLf & HAMU_CommandIntro("HAMU_FilterRowsToNewTable") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_SampleData") & vbCrLf & HAMU_CommandIntro("HAMU_SampleData") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_ShuffleRows") & vbCrLf & HAMU_CommandIntro("HAMU_ShuffleRows") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_CompareTwoLists") & vbCrLf & HAMU_CommandIntro("HAMU_CompareTwoLists") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_ExtractUniqueList") & vbCrLf & HAMU_CommandIntro("HAMU_ExtractUniqueList") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_DataProfile") & vbCrLf & HAMU_CommandIntro("HAMU_DataProfile") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_DataQualityReport") & vbCrLf & HAMU_CommandIntro("HAMU_DataQualityReport") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_FindMissingCombinations") & vbCrLf & HAMU_CommandIntro("HAMU_FindMissingCombinations") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_CleanData") & vbCrLf & HAMU_CommandIntro("HAMU_CleanData") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_TextCase") & vbCrLf & HAMU_CommandIntro("HAMU_TextCase") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_PatternFind") & vbCrLf & HAMU_CommandIntro("HAMU_PatternFind") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_RegexReplace") & vbCrLf & HAMU_CommandIntro("HAMU_RegexReplace") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_ExtractDigits") & vbCrLf & HAMU_CommandIntro("HAMU_ExtractDigits") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_ConvertQuotedNumbers") & vbCrLf & HAMU_CommandIntro("HAMU_ConvertQuotedNumbers") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_MaskData") & vbCrLf & HAMU_CommandIntro("HAMU_MaskData") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_ConvertFormulasToValues") & vbCrLf & HAMU_CommandIntro("HAMU_ConvertFormulasToValues") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_UnmergeAndFill") & vbCrLf & HAMU_CommandIntro("HAMU_UnmergeAndFill") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_StandardFormat") & vbCrLf & HAMU_CommandIntro("HAMU_StandardFormat") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_DeleteBlankCells") & vbCrLf & HAMU_CommandIntro("HAMU_DeleteBlankCells") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_DeleteDropdowns") & vbCrLf & HAMU_CommandIntro("HAMU_DeleteDropdowns") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_RowStriping") & vbCrLf & HAMU_CommandIntro("HAMU_RowStriping") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_DatePicker") & vbCrLf & HAMU_CommandIntro("HAMU_DatePicker") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_FixDates") & vbCrLf & HAMU_CommandIntro("HAMU_FixDates") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_CheckDates") & vbCrLf & HAMU_CommandIntro("HAMU_CheckDates") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_DateBuilder") & vbCrLf & HAMU_CommandIntro("HAMU_DateBuilder") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_SheetIndex") & vbCrLf & HAMU_CommandIntro("HAMU_SheetIndex") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_SortSheets") & vbCrLf & HAMU_CommandIntro("HAMU_SortSheets") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_ShowHidden") & vbCrLf & HAMU_CommandIntro("HAMU_ShowHidden") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_DynamicValidation") & vbCrLf & HAMU_CommandIntro("HAMU_DynamicValidation") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_SelectionToPNG") & vbCrLf & HAMU_CommandIntro("HAMU_SelectionToPNG") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_InsertPhotosFromFolder") & vbCrLf & HAMU_CommandIntro("HAMU_InsertPhotosFromFolder") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_CreateQRCode") & vbCrLf & HAMU_CommandIntro("HAMU_CreateQRCode") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_CompareSheets") & vbCrLf & HAMU_CommandIntro("HAMU_CompareSheets") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_MarkProblemCells") & vbCrLf & HAMU_CommandIntro("HAMU_MarkProblemCells") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_HighlightByCriteria") & vbCrLf & HAMU_CommandIntro("HAMU_HighlightByCriteria") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_HighlightDuplicates") & vbCrLf & HAMU_CommandIntro("HAMU_HighlightDuplicates") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_AddDuplicateCount") & vbCrLf & HAMU_CommandIntro("HAMU_AddDuplicateCount") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_ListLinks") & vbCrLf & HAMU_CommandIntro("HAMU_ListLinks") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_BreakLinks") & vbCrLf & HAMU_CommandIntro("HAMU_BreakLinks") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_Help") & vbCrLf & HAMU_CommandIntro("HAMU_Help") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_About") & vbCrLf & HAMU_CommandIntro("HAMU_About") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_AutoFit") & vbCrLf & HAMU_CommandIntro("HAMU_AutoFit") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_AutoTable") & vbCrLf & HAMU_CommandIntro("HAMU_AutoTable") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_ResizeTable") & vbCrLf & HAMU_CommandIntro("HAMU_ResizeTable") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_UDFHelp") & vbCrLf & HAMU_CommandIntro("HAMU_UDFHelp") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("UDF_URLDecode") & vbCrLf & HAMU_CommandIntro("UDF_URLDecode") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("UDF_RenklileriTopla") & vbCrLf & HAMU_CommandIntro("UDF_RenklileriTopla") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("UDF_RengeGoreSay") & vbCrLf & HAMU_CommandIntro("UDF_RengeGoreSay") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("UDF_YerelDosyaYolu") & vbCrLf & HAMU_CommandIntro("UDF_YerelDosyaYolu") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("UDF_GizliLink") & vbCrLf & HAMU_CommandIntro("UDF_GizliLink") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("UDF_LinkParamDegeri") & vbCrLf & HAMU_CommandIntro("UDF_LinkParamDegeri") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("UDF_SayidanMetine") & vbCrLf & HAMU_CommandIntro("UDF_SayidanMetine") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("UDF_TablodanVeriGetir") & vbCrLf & HAMU_CommandIntro("UDF_TablodanVeriGetir") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_CountByColor") & vbCrLf & HAMU_CommandIntro("HAMU_CountByColor") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_SumByColor") & vbCrLf & HAMU_CommandIntro("HAMU_SumByColor") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_FillBlanksZero") & vbCrLf & HAMU_CommandIntro("HAMU_FillBlanksZero") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_AddTextAffixes") & vbCrLf & HAMU_CommandIntro("HAMU_AddTextAffixes") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_LeadingZeros") & vbCrLf & HAMU_CommandIntro("HAMU_LeadingZeros") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_DeduplicateItems") & vbCrLf & HAMU_CommandIntro("HAMU_DeduplicateItems") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_ExtractTextPart") & vbCrLf & HAMU_CommandIntro("HAMU_ExtractTextPart") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_DateRangeList") & vbCrLf & HAMU_CommandIntro("HAMU_DateRangeList") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_FormulaAudit") & vbCrLf & HAMU_CommandIntro("HAMU_FormulaAudit") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_SelectSpecialCells") & vbCrLf & HAMU_CommandIntro("HAMU_SelectSpecialCells") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_ListDefinedNames") & vbCrLf & HAMU_CommandIntro("HAMU_ListDefinedNames") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_DataEntry") & vbCrLf & HAMU_CommandIntro("HAMU_DataEntry") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_ClearHighlights") & vbCrLf & HAMU_CommandIntro("HAMU_ClearHighlights") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("UDF_LambdaMultiLookup") & vbCrLf & HAMU_CommandIntro("UDF_LambdaMultiLookup") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("UDF_LambdaMultiSum") & vbCrLf & HAMU_CommandIntro("UDF_LambdaMultiSum") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("UDF_LambdaUniqueMerge") & vbCrLf & HAMU_CommandIntro("UDF_LambdaUniqueMerge") & vbCrLf & vbCrLf
 HAMU_CatalogGuide = HAMU_CatalogGuide & HAMU_CommandLabel("HAMU_SaveCopy") & vbCrLf & HAMU_CommandIntro("HAMU_SaveCopy") & vbCrLf & vbCrLf
End Function

Private Function HAMU_FilesGuide() As String
 Dim tr As String, en As String
 tr = tr & "Yedek Oluştur, mevcut dosyanın tarih/saatli kopyasını Ayarlar bölümündeki klasöre kaydeder. Boş yol Belgeler/HAMU/Backup"
 tr = tr & (" demektir. Sonuç ekranındaki Klasörü Aç düğmesi konumu açar." & vbLf & vbLf & "Tarihli Sürüm, kayıtlı dosyanın aynı klasörde tarih/saatli")
 tr = tr & " yeni kopyasını açıp doğrular; başarılıysa eski dosyayı kapatır ve siler. Kayıtsız veya salt okunur dosyada önce kayıt/y"
 tr = tr & ("azma izni gerekir." & vbLf & vbLf & "Klasör İçeriği, mevcut klasörü veya alt klasörlerle birlikte dosyaları listeler: bağlantı, yerel yol")
 tr = tr & (", uzantı/tür, okunabilir boyut ve tarihler. Kapsam Ayarlar’dan sabitlenebilir." & vbLf & vbLf & "PDF Kaydet, seçili aralık veya etkin say")
 tr = tr & "fa için çalışır. Otomatik yön genişlik/yükseklik oranıyla belirlenir; dikey/yatay elle seçilebilir. Genişlik bir sayfaya"
 tr = tr & (" sığar, uzun listeler aşağıya devam eder. Baskı alanı ve yön işlem sonunda eski durumuna döner." & vbLf & vbLf & "Aktarma Sihirbazı, LAMB")
 tr = tr & "DA tanımları, özel hücre stilleri, adlandırılmış veri aralıkları, sayfa verisi ve Excel tablolarını seçerek aktarır. Say"
 tr = tr & "fa/tablo/aralık aktarımları değer ve biçim kopyalarıdır; VBA veya kaynak hücre formülleri taşınmaz. Şablon salt okunur a"
 tr = tr & "çılır. Var olan sayfa silinmez; Atla veya Yeniden Adlandır seçin. Önek verilen LAMBDA bağımlılıklarını hedef kitapta kon"
 tr = tr & "trol edin."
 en = en & "Create Backup saves a timestamped copy to the folder in Settings. An empty path uses Documents/HAMU/Backup. Open Folder "
 en = en & ("in the result opens that location." & vbLf & vbLf & "Dated Version opens and verifies a new timestamped copy in the original folder, then")
 en = en & (" closes and deletes the original after success. Save unsaved workbooks first; read-only files cannot be replaced." & vbLf & vbLf & "Folde")
 en = en & "r Contents lists the current folder or subfolders: link, local path, extension/type, readable size and timestamps. You c"
 en = en & ("an preset the scope in Settings." & vbLf & vbLf & "Save PDF exports the selection or active sheet. Automatic orientation uses the width/h")
 en = en & "eight ratio; portrait/landscape can be chosen manually. Width fits one page; long lists continue downward. Prior print a"
 en = en & ("rea and orientation are restored." & vbLf & vbLf & "Transfer Wizard transfers selected LAMBDA definitions, custom cell styles, named data")
 en = en & " ranges, sheet data and Excel tables. Sheet/table/range snapshots contain values and formatting, without VBA or original"
 en = en & " cell formulas. The template opens read-only. Existing sheets are preserved; choose Skip or Rename. Review dependencies "
 en = en & "when adding prefixes to LAMBDA names."
 HAMU_FilesGuide = HAMU_L(tr, en)
End Function


