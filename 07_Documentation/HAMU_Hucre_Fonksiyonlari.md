# HAMU Tools V1 — Hücre Fonksiyonları

Excel’de Fonksiyon Ekle penceresinin **HAMU Tools** kategorisinde bulunurlar. Yardım ve Hakkında > Fonksiyonlar menüsündeki her başlık örnek ve açıklama gösterir. İç yardımcı modüller Option Private Module ile listeden gizlenir.

## HAMU_URLCoz

=HAMU_URLCoz(A1) | UTF-8 URL karakterlerini çözer. Geçersiz yüzde dizisini korur.

## HAMU_RengeGoreTopla

=HAMU_RengeGoreTopla(A1:A20;B1) | B1 ile aynı dolgu rengine sahip sayıları toplar. Biçim değişimi sonrası Ctrl+Alt+F9 kullanın.

## HAMU_RengeGoreSay

=HAMU_RengeGoreSay(A1:A20;B1) | B1 ile aynı dolgu rengine sahip hücreleri sayar. Biçim değişimi sonrası Ctrl+Alt+F9 kullanın.

## HAMU_YerelDosyaYolu

=HAMU_YerelDosyaYolu(A1) | Kişisel OneDrive URL yolunu yerel OneDrive ortam değişkeniyle eşler; diğer URL türlerini korur.

## HAMU_GizliLink

=HAMU_GizliLink(A1) | Hücrenin ilk köprü adresini getirir.

## HAMU_LinkParamDegeri

=HAMU_LinkParamDegeri(A1;"id") | URL içindeki sorgu parametresinin değerini getirir.

## HAMU_SayidanMetine

=HAMU_SayidanMetine(123) | Türkçe: yüz yirmi üç. 15 basamaktan küçük mutlak değerler; ondalık kısmı tek tek okur.

## HAMU_TablodanVeriGetir

=HAMU_TablodanVeriGetir(A1;"Tablo1";"Kod";"Ad") | Formülün bulunduğu kitapta tablodan ilk eşleşen değeri getirir.

## Renk hesabı

Bu fonksiyonlar doğrudan hücre dolgu rengini karşılaştırır; koşullu biçimlendirme rengini kullanmaz. Dolgu rengi değiştiğinde Ctrl+Alt+F9 ile yeniden hesaplayın. Örnek renk hücresinin ilk hücresi kullanılır.

Eski kısa fonksiyon adları yalnız gizli uyumluluk modülünde tutulur. Yeni formüllerde HAMU_ adlarını kullanın.

Geliştirici kaynakları: [MacroOptions](https://learn.microsoft.com/en-us/office/vba/api/excel.application.macrooptions), [Option Private Module](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-private-statement).

Renk fonksiyonları 250.000 hücreyle sınırlıdır ve volatile değildir. Tüm sütun yerine gerçek veri aralığını seçin. Yalnız biçim değişikliklerinde Ctrl+Alt+F9 gerekir.

## Gerçek Excel LAMBDA tanımları



**Bilgi > LAMBDA Şablonu > Şablonu Aç** yolundan `09_Templates/HAMU_LAMBDA_Sablonu.xlsx` açılır. Aktarma Sihirbazı ile tanımları çalışma kitabınıza aktarabilirsiniz. Bunlar VBA yardımcı fonksiyonu değildir; Excel'in Ad Yöneticisi'nde görünür.

1. `HAMU_L_CokluAra`: iki anahtar/sonuç aralığı çiftinde sırayla tam eşleşme arar; ilk sonucu verir. Bulamazsa boş döner. Tek sütunlu ve eşit uzunlukta eşleşme çiftleri gerekir.
2. `HAMU_L_CokluTopla`: iki anahtar/değer aralığındaki aynı anahtarın sayısal karşılıklarını toplar.
3. `HAMU_L_BenzersizBirlestir`: iki aralığı tek listeye birleştirir, boşları kaldırır, benzersiz değerleri sıralar.

Örnekler şablonda hesaplanmış halde bulunur. Bu şablon LAMBDA/VSTACK/TOCOL gibi güncel Excel işlevlerini gerektirir; güncel Microsoft 365 veya Excel 2024 kullanın. Eski Excel sürümlerinde çalışmayabilir. Excel'in kendi gizli `_xlfn` / `_xlpm` uyumluluk adları silinmez. Eklentinin iç yardımcı modülleri Option Private Module ile gizlidir; sekiz HAMU VBA hücre fonksiyonu kullanılmaya devam eder.

