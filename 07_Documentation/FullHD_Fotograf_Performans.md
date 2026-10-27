# Toplu fotoğraf: Full HD ve performans kontrolü

27.10.2026. Başlangıç dosyası: AddIns/HAMUToolsV1.xlam (SHA256 c9fc0e3495cb2e9b6dcf23f154bad0827eb03fc145386026fa3602bf2e797b13).

## Değişiklik

Klasör seçiminden sonra tek pencerede yerleşim (hücre içine / hücre üstüne) ve kalite (Full HD optimize / orijinal kalite) seçilir. Full HD varsayılandır. Yatay görseller en fazla 1920 × 1080, dikey görseller 1080 × 1920 piksel sınırına oranı korunarak sığdırılır. Küçük fotoğraflar büyütülmez. Saydam görseller PNG, diğer görseller JPEG kalite 85 olarak hazırlanır. Küçük JPEG/PNG dosyalarında yeniden kodlama boyutu artırıyorsa özgün dosya kullanılır. EXIF yönü optimizasyon sırasında uygulanır. Kaynak dosyalara yazılmaz.

Fotoğraf kaynakları yalnızca komut çalıştığında okunur. Geçici betikler okunabilir UTF-8 metin olarak yazılır; EncodedCommand kaldırıldı. Optimizasyon yardımcı kodu orijinal kalite işleminde yüklenmez. Hücre içi aktarım 50 fotoğraflık bloklarla çalışır. Geçici dosyalar başarı/hata sonunda temizlenir. Hücre üstüne eklemede, ölçüleri WIA ile okunabilen dosyalar için tam boyutlu ara Excel resim nesnesi oluşturulmaz. İşlem sonundaki Excel ayarları açıkça geri yüklenir.

## Ölçüm

Altı görsel: büyük yatay JPEG, aynı kaynakla ikinci fotoğraf, dikey JPEG, saydam PNG, küçük JPEG ve EXIF yön bilgili JPEG. Kaydetme sonrasında Excel aynı resimleri tekilleştirebildiği için paket beş medya parçası içerir. Ölçümler tek çalıştırmadır; kullanıcı dosyalarına göre değişir.

| Yerleşim / kalite | Dosya boyutu | Ekleme | Kaydetme | Yeniden açma |
|---|---:|---:|---:|---:|
| Hücre içine / orijinal | 32.07 MB | 3.54 sn | 28.84 sn | 1.00 sn |
| Hücre içine / Full HD | 1.57 MB | 2.65 sn | 3.53 sn | 0.45 sn |
| Hücre üstüne / orijinal | 32.07 MB | 0.15 sn | 20.53 sn | 0.55 sn |
| Hücre üstüne / Full HD | 1.57 MB | 1.84 sn | 3.20 sn | 0.41 sn |

Dosya boyutu yaklaşık %95 azaldı. Optimizasyonun hazırlık maliyeti vardır: hücre üstüne eklemenin kendisi orijinal moddan daha yavaş olabilir; dosyanın kaydı ve yeniden açılması belirgin şekilde hafifledi. Orijinal kaliteyi seçmek büyük dosya ve bellek maliyetini korur.

## Doğrulama

- İki yerleşim ve iki kalite yolu çalıştırıldı. Hücre içi fotoğraflarda dışarıda nesne kalmadı; hücre üstü fotoğraflar hedef hücre sınırlarına sığdı.
- Paket içindeki optimize görsellerin piksel boyutları doğrulandı; saydamlık kanalı kaldı. 320 × 180 küçük görsel büyütülmedi.
- PNG/JPEG/WebP/BMP/GIF/TIFF optimizasyonu ve 51 fotoğraflı iki bloklu aktarım geçti.
- Korumalı hedef ve okunamayan dosya testleri geçti. Çalışma kitabı, geçici dosya ve uygulama ayarı temizliği kontrol edildi.
- Yeni form görsel olarak incelendi. Ribbon XML ve ikon parçaları önceki dosyayla birebir aynı kaldı. Diğer form tasarımları değiştirilmedi.

## Boştayken çalışma ve kapsam

Kaynaklarda Application düzeyinde WithEvents, OnTime, OnKey, Workbook_Open, Auto_Open, SheetChange, SelectionChange veya Application.Volatile çağrısı bulunmadı. Takvimdeki WithEvents yalnızca açık formdaki gün etiketlerine aittir. Eklenti kaynak sayfasında işlem başlatan formül yoktur.

Ayrı Excel sürecinde ölçüldü. Eklentisiz son iki 3 saniyelik pencerede CPU zamanı 0.031/0.031 saniye; eklenti yüklüyken 0.063/0.047/0.031 saniye. Eklentili ölçüm sırasında çalışma kümesi yaklaşık 256.1 MB'de sabit kaldı (eklentisiz yaklaşık 229.9 MB). İlk eklentisiz pencerede Office başlangıç faaliyeti sürdüğü için 1.56 saniye CPU görüldü; bunu eklenti karşılaştırmasında kazanç olarak yorumlamıyoruz. Bu kısa testte sürekli yüksek CPU veya büyüyen bellek gözlenmedi; tüm Excel donmalarının giderildiği iddia edilmez. Excel, kapanan büyük resimli kitaplardan sonra kendi önbelleğini bir süre tutabilir.

## Kaynakları düzenleme

VBA: 02_Modules/Features/modHAMU_PhotoBatch.bas ve modHAMU_Visual.bas.
Form: 05_Forms/frmHAMU_PhotoOptions.frm/.frx.
Betik: 09_Templates/PhotoBatch/photo-optimize.ps1.
Eklentide HAMU_PhotoResources çok gizli sayfasının A1 hücresi XLSX şablonunu, A2 hücresi paketleme betiğini, A3 hücresi optimizasyon betiğini içerir. Betik kaynakları değiştirilirse ilgili hücreler de güncellenmelidir. Başlangıçta bu betikler çalıştırılmaz.

Final sürüm AddIns/HAMUToolsV1.xlam yoluna yedek alınarak kuruldu. Normal Excel başlangıcında boş test kitabı açıldı; yükleme uyarısı görülmedi ve HAMU Tools sekmesi seçilerek komutlar görsel olarak doğrulandı.

Optimizasyon yeni eklenen fotoğraflara uygulanır; daha önce dosyaya eklenmiş büyük resimleri otomatik dönüştürmez.

Final sürüm kullanıcı tarafından kontrol edilerek onaylandı.
