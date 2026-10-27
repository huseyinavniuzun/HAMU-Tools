# Fotoğraf modülü yükleme düzeltmesi

27.10.2026

Önceki sürüm, otomasyonla yükleme testlerini geçmesine rağmen Excel’in normal başlangıcında modHAMU_PhotoBatch yükleme uyarısını verdi. Bu davranış normal Excel başlangıcında yeniden üretildi. Yalnızca otomasyonla yapılan açma testi yeterli değildi.

Fotoğraf modülündeki uzun gömülü XLSX şablonu ve yardımcı PowerShell metni koddan çıkarıldı. Eklentinin HAMU_PhotoResources adlı çok gizli sayfasında A1 şablonun Base64 verisini, A2 okunabilir betik metnini tutar. Bunlar yalnızca fotoğraf komutu çağrıldığında okunur; başlangıç olayı veya arka plan dinleyicisi eklenmedi. Kaynakların düzenlenebilir halleri 09_Templates/PhotoBatch klasöründedir. VBA kaynakları Windows-1254, BOM olmadan saklanır.

Bu değişiklikten sonra test adlı eklenti normal başlangıçta yüklendi ve 115 satırlık fotoğraf modülü eksiksiz okunabildi. Kalıcı HAMUToolsV1.xlam adıyla normal başlangıçta da hata penceresi gözlenmedi ve HAMU sekmesi göründü. Kalıcı dosya AppData/Roaming/Microsoft/AddIns/HAMUToolsV1.xlam konumundadır. OPEN kaydında yalnızca HAMUToolsV1.xlam kaldı.

Üç biçimle toplu fotoğraf testi geçti (PNG, JPEG, WebP); hücre dışı nesne kalmadı ve uygulama ayarları geri yüklendi. PNG ve JPEG dosya baytları çıktı paketinde aynen korundu. WebP için Excel paketleme sırasında biçim dönüşümü yapabilir; bu rapor WebP için bayt eşitliği iddiasında bulunmaz.

Not: Sorunu hangi VBA/Office iç mekanizmasının tetiklediği kesinleşmedi. Kanıt, normal başlangıçta önce/sonra davranışı ve kaynak ayrımının sonucudur. Normal başlangıç kontrolü tamamlandı; kullanıcı final sürümü kontrol ederek onayladı.
