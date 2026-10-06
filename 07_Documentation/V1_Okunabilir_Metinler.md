# V1 — Okunabilir metinleri düzenleme

6.097 hex metin çağrısı okunabilir metne dönüştürüldü. Türkçe karakterler doğrudan yazılıdır; hex çözme veya hash anahtarı hesaplama gerekmez.

## Nereden düzenlenir?

- `02_Modules/Core/modHAMU_UIText.bas`: Ribbon ve Dosya merkezi başlıkları.
- `02_Modules/Core/modHAMU_CommandInfo.bas`: komut adları ve açıklamaları.
- `02_Modules/Core/modHAMU_HelpContent.bas`: kullanım rehberi.
- `02_Modules/Core/modHAMU_Language.bas`: 1.196 okunabilir Türkçe/İngilizce çeviri çifti.
- `05_Forms/*.frm`: form kodları. Form öğelerinin yerleşimi Excel VBA düzenleyicisindeki UserForm tasarım ekranında değiştirilir; FRX dosyası metin dosyası değildir.

Örnekler:

```vb
HAMU_L("Hızlı" & vbLf & "Tablo", "Quick" & vbLf & "Table")
HAMU_Text("İşlem tamamlandı.")
mEnglish.Add "İşlem tamamlandı.", "Operation completed."
```

`HAMU_Text` artık hex çözmez; doğrudan metni çeviri sözlüğünde arar. İngilizce karşılık bulunmazsa metni olduğu gibi gösterir. Sözlükteki Türkçe kaynak metni ve kullanılan metni birlikte değiştirin. `HAMU_L` içindeki iki dil karşılığı aynı satırda düzenlenebilir.

## Dosya kodlaması

VBA kaynakları `.bas`, `.frm`, `.cls` **Windows-1254 (Türkçe), BOM olmadan** kaydedilir. Bunları UTF-8 olarak kaydedip doğrudan VBA'ya içe aktarmayın. XML, JSON ve Markdown dosyaları UTF-8'dir. Türkçe harfler okunabilir metindir; CP1254 dışında kalan ok simgesi yalnız `ChrW(8594)` ile yazılır.

XLAM içindeki kodlar da okunabilirdir. Önce yedek alın; Excel VBA düzenleyicisinden kodu değiştirip eklentiyi kaydedebilirsiniz. Dışarıdaki kaynak dosyasını değiştirmek hazır XLAM'ı kendiliğinden güncellemez; ilgili modülü içe aktarmak veya eklentiyi yeniden oluşturmak gerekir. JSON metin envanterleri çalışma sırasında yüklenmez.

RibbonX Editor ile XML düzenlemesi için önceki RibbonX düzenleme rehberini kullanın. Dinamik TR/EN başlıklar VBA callback'lerinden gelir; yalnız Türkçe sabit başlık isteyenler için düzenlenebilir TR XML ayrıca bulunur.

## Kontrol

Gerçek Excel oturumunda bir başlık elle `Elle İĞŞıçöü` olarak değiştirildi; Türkçe görünümü, İngilizce karşılığı ve yeni sözlük kaydı doğrulandı. Değişiklik test kopyasında yapıldı. Derleme, dil geçişi, iş akışları ve kullanım güvenliği testleri geçti. Canlı Ribbon çizimi ve farklı DPI ölçekleri bu metin düzenlemesinde ayrıca görsel olarak test edilmedi.
