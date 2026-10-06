# Yeni özellik ekleme rehberi

## Mimari

İş mantığı Features modüllerinde, ortak davranışlar Core modüllerindedir. `modHAMU_Ribbon.HAMU_OnAction`, Ribbon kontrolünün `Tag` değerini `HAMU_RunCommand` girişine iletir. Bu giriş, açık çalışma kitabı kontrolünden sonra ilgili işlemi çağırır. `commands.json` teslim edilen araçların ad, kategori, alt kategori ve açıklama envanteridir.

## Yeni bir işlem

1. İlgili Features modülüne `Public Sub HAMU_YeniIslem()` ekleyin. Yeni bir konuysa ayrı `.bas` oluşturun.
2. Aralık için `HAMU_PromptRange` kullanın ve sonuç `Nothing` ise çıkın. Seçimi doğrudan `Selection` üzerinden değiştirmeyin.
3. Metin seçeneği için `HAMU_InputText`; sayı için `HAMU_InputValue(..., ValueType:=1)` kullanın. Sayı iptalinde Boolean False döner: `VarType(sonuc) = vbBoolean` ile ayırın; sayı 0 ile iptali karıştırmayın.
4. Yazma, sınır ve veri yapısı kontrollerini işlemin kendi kodunda yapın. Tek bir Range varsayan araçlarda çok alanlı seçimi kabul etmeyin.
5. `clsHAMU_AppState` oluşturup `Capture` çağırın. Yordam sonlandığında sınıf eski ayarları geri yükler; gerekirse açık `Restore` çağırın.
6. `HAMU_RunCommand` içindeki Select Case'e yeni Tag / yordam eşlemesini ekleyin.
7. `commands.json` envanterine ekleyin; Custom UI XML'de ilgili Ribbon ve Backstage kategorilerine aynı Tag ile kontrol ekleyin. `screentip/supertip` ve Backstage `description` doldurun.
8. İkona anlamlı ASCII bir ad verin; 16/32/64 PNG ve master kaynaklarını ekleyin. XML image adı görsel ilişkisinin Id'siyle aynı olmalıdır. Master'ı Custom UI içine koymayın.
9. Kaynakları Excel'e aktarın, ThisWorkbook kodunu mevcut kitap nesnesine yerleştirin, Debug > Compile VBAProject çalıştırın. `.cls` dosyasını sıradan bir sınıf olarak import etmek çalışma kitabı olaylarının yerine geçmez.
10. XLAM içine güncel Custom UI ve görsel ilişkilerini bir Open XML / Ribbon düzenleyicisiyle yeniden ekleyin. Sadece `.bas` dosyasını diskte değiştirmek hazır `.xlam` içindeki kodu değiştirmez.

## Kodlama ve kaynak içe aktarma

VBA dışa aktarımları Türkçe Windows ANSI / CP1254'tür; XML, JSON ve belgeler UTF-8'dir. Türkçe harfleri koruyan bir düzenleyici kullanın. `.frm` ve `.frx` aynı klasörde ve aynı temel adda tutulmalıdır. FRX gerçek ikili form kaynağıdır; metin olarak düzenlemeyin.

## Profil ve görünüm

HAMU adı, web adresi ve sürüm `Core/modHAMU_Constants.bas` içinde tek yerden değiştirilir. Ortak tema `Core/modHAMU_Forms.bas` içindedir. Form kontrol adlarını değiştirmek ilgili olay yordamlarını da değiştirmeyi gerektirir. İkonlarda viewBox 32×32, açık gri yüzey ve renkli işlev nesneleri kullanılır. Güncel özgün logo master/hamu512_original.png dosyasıdır.


## Komut bilgileri ve hata akışı

Yeni komutu commands.json'a, modHAMU_Ribbon.HAMU_RunCommand yönlendirmesine ve iki Custom UI yüzeyine ekleyin. modHAMU_CommandInfo'daki ad/açıklama/sayfa gereksinimi eşleşmesini güncelleyin. Form açıklamaları gHAMU_CommandId ile bu ortak kaynağı kullanır. Türkçe VBA metinleri doğrudan okunabilir metin olarak yazılır; kod dosyaları BOM olmadan CP1254'tür. Kullanıcı komutlarında kitap/sayfa giriş kontrolü ve HAMU_ShowError kullanın; yardımcılar beklenen hatayı çağırana iletsin. Dosya okurken HAMU_OpenSourceWorkbook ve owned bayrağıyla yalnız size ait oturumu kapatın.


## Dil ve ayar eklemek

`modHAMU_Settings` kullanıcı tercihlerini yükler/kaydeder. Yeni tercih için varsayılan anahtar, doğrulama ve Ayarlar formundaki giriş birlikte eklenmelidir. Kitap açma/kaydetme olaylarını genel Excel uygulamasına bağlamayın.

Custom UI metinleri `getLabel/getDescription/getHelperText/getScreentip/getSupertip/getAltText` callback'leriyle `modHAMU_UIText` üzerinden alınır; kaynak karşılıkları UI_Metinleri.json dosyasındadır. Dil_Kaynaklari.json çeviri envanteridir. Dil kaynaklarının değişmesi VBA modüllerinin ve XLAM'ın yeniden derlenmesini gerektirir. `HAMU_L(turkce, english)` kısa yeni metinlerde; `HAMU_Text("Türkçe metin")` okunabilir metni çeviri sözlüğünde arar; hex kullanılmaz. Regex, formül, dosya yolu ve kullanıcı girdilerini çevirmeyin. Türkçe çalışma kitabı kod adını ThisWorkbook olarak yeniden adlandırmayın; kodu mevcut kitap bileşenine yerleştirin.

Aralık için mevcut seçimi kullanabilen araçlarda `HAMU_WorkRange`, açık hedef isteyenlerde `HAMU_PromptRange` kullanın. İptalde Nothing ile hemen çıkın; başka bir modal form açıkken ikinci bir modal aralık seçici açmayın. Durum çubuğu dönüşünde Boolean False yerine `clsHAMU_AppState.Restore` kullanın.

## V1 aralık ve çıktı sözleşmeleri

HAMU_BoundedRange(source, trimUsedExtent=False) bütün satır/sütun seçimini gerçek formül/değer sonuna daraltır. trimUsedExtent=True, UsedRange kaynaklarındaki yalnız biçim içeren uzak hücreleri de çıkarır. Açıkça seçilmiş kısmi aralıklar varsayılan olarak korunur. Döngü başlamadan CountLarge sınırı uygulanır; çıktı adresini veri kaynağı gibi daraltmayın. Range.Find çağrılarında tüm arama seçeneklerini açıkça belirtin.

HAMU_ExportRangeImage seçim ve uygulama durumunu geri yükler, geçici grafiği siler. Clipboard işlemleri yeniden girişe karşı ortak dispatcher koruması altında çalışır. Dosya varlığının yanında içerik pikselleri de test edilir. PDF yardımcı metodu PageSetup ayarlarını geri yükler.

LAMBDA tanımları 09_Templates kitabındadır; modHAMU_UDF sadece sekiz VBA hücre işlevini dışarı açar. Şablondaki gizli Excel uyumluluk adlarını gereksiz kullanıcı işlevi sanıp silmeyin. Aktarım testinde hedef kitabın görünür adlarını sayın.
