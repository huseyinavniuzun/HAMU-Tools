# HAMU Tools V1 — Ribbon yerleşimi ve RibbonX düzenleme

- Hızlı / Tablo düğmesi iki satırdır.
- Karekod / Barkod tek düğme ve tek formdur. Kod Türü seçiminde Karekod ve Barkod bulunur; model listesi türe göre güncellenir. İçerik, içerik türü ve kaynak/hedef adresi korunur. Toplu giriş aynı ekrandadır.
- Sayfa İndeksi ve Toplu Fotoğraf, Sayfa ve Görsel grubunun başında büyük düğmelerdir.
- Sayfaları Karşılaştır, Veri > Ayıkla menüsünde sondan ikinci sıradadır.
- Sorunlu / Hücreler düğmesi, Denetim / Araçları menüsünün önündedir.
- Bilgi grubunda fx simgeli büyük Ek Fonksiyon menüsü vardır. Hücre Fonksiyonları ve LAMBDA Şablonları bu menüdedir.
- Dosya > HAMU Tools'un iki sütununda da bir boş satırla başlangıç yapılır. Office sürümüne göre gerçek çizim yüksekliği değişebilir.

## RibbonX Editor hatası

[Content_Types].xml kökü önce ns0:Types biçimindeydi. Excel bu XML'i okuyabiliyor, ancak WindowsBase paket okuyucusu kullanıcının bildirdiği Required Types / Gerekli Types etiketi bulunamadı hatasını veriyordu. Kök artık varsayılan içerik-türü ad alanında Types olarak yazılır. Paket ilişkileri de varsayılan ad alanıyla yazılır. Düzeltme paketleyicide kalıcıdır.

Eski pakette hata aynı WindowsBase API'sinde yeniden üretildi. Yeni paket açıldı; Custom UI parçasına değişiklik yazılıp kaydedildi; kaydedilmiş paket yeniden açıldı. VBA proje içeriğinin değişmediği doğrulandı. RibbonX Editor grafik arayüzünde elle açma sınanmadı; hatayı veren paket okuyucusunun açma/yazma yolu doğrudan sınandı.

## Düzenleme

1. Excel'i tamamen kapatın. 08_Addin/HAMUToolsV1.xlam dosyasının bir kopyasını alın ve Office RibbonX Editor ile açın.
2. Office 2010+ Custom UI bölümünde grupları, menüleri ve düğmeleri taşıyabilirsiniz. Değişiklikten sonra doğrulayıp kaydedin.
3. Standart XML'de başlıklar TR/EN desteği için getLabel callback'inden gelir. Adı XML içinde değiştirmek için ilgili kontrolün getLabel niteliğini kaldırıp label="Yeni ad" yazın. Aynı kontrol üzerinde label ve getLabel birlikte olamaz.
4. Türkçe başlıkları doğrudan içeren 01_CustomUI/customUI14_duzenlenebilir_TR.xml dosyası da verildi. Editor'daki Custom UI içeriğini bu dosyanın içeriğiyle değiştirirseniz başlıkları doğrudan düzenleyebilirsiniz. Bu sürüm başlıkları Türkçe sabitler; İngilizce ayarı bu statik başlıkları değiştirmez. Standart customUI14.xml TR/EN desteğini korur.
5. Simge eklemek için Editor üzerinden görseli ekleyin; image niteliğini onun adına bağlayın. Mevcut simgeler XLAM'a gömülüdür. Yeni işlemin onAction/tag bağlantısına karşılık gelen VBA kodu gerekir.
6. Excel'i yeniden açın. RibbonX kaynak düzenlemesi, Excel Seçenekleri > Şeridi Özelleştir ekranındaki kişisel düzenlemeden farklıdır.

93 menü kaydı, dört veri menüsü, 15 form çifti; tam derleme, XSD, birleşik tür/model geçişi ve veri/adres koruması, ortak iş akışları ve kitap yokken 79 komut testleri geçti. Ayrıntılar Test_Ciktisi.txt ve Validation.json içindedir.
