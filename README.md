# PrintBatch

Windows için sade toplu belge yazdırma uygulaması. PDF, DOC, DOCX, XLS, XLSX, PPT ve PPTX desteklenir.

## Çalıştırma

İlk kurulumda proje klasöründe `python setup_dependencies.py` komutunu çalıştırın. Bu komut, doğrulanmış SumatraPDF sürümünü indirir.

**`Baslat.vbs` dosyasına çift tıklayın.** Bu bilgisayarda bulunan Python'u otomatik kullanır. `printer_app.py`, `office_backend.ps1` ve `tools` klasörü aynı yerde kalmalıdır.

Alternatif: `python printer_app.py`. Python 3.10+ ve Tkinter gerekir; uygulama harici Python paketi kullanmaz. Office belgelerinde kurulu LibreOffice öncelikle kullanılır; bulunamazsa ilgili Microsoft Word, Excel veya PowerPoint uygulaması kullanılır. PDF için ikisi de gerekmez. Yazıcı listesini okumak ve Office yedek dönüşümü için Windows PowerShell'in yerel betikleri çalıştırmasına izin verilen normal kullanıcı oturumu gerekir; yönetici izni gerekmez. Uygulama kullanıcının güvenlik ayarlarını değiştirmez.

1. Dosyaları ekleyin veya pencereye sürükleyin.
2. Yazıcıyı, kopya sayısını ve diğer ayarları seçin.
3. Bir belgeyi seçip **Seçileni önizle** ile kontrol edin.
4. **Tümünü yazdır** düğmesine basın.

Sayfa aralığı boşsa tüm sayfalar yazdırılır; örnek: `1-3,5`. Ayarlar listedeki tüm belgelere uygulanır. Satıra çift tıklayarak uzun hata mesajını okuyabilirsiniz. Hatalı belge atlanır ve sonraki belgeye geçilir.

Office belgeleri mevcut sayfa düzenleriyle geçici PDF’ye dönüştürülür; Excel'in kayıtlı baskı alanları korunur. LibreOffice, Microsoft Office'ten farklı sayfalama yapabileceğinden önemli belgeleri önizlemede kontrol edin. Belgeler internete yüklenmez, kaynak dosyalar değiştirilmez, makrolar çalıştırılmaz. LibreOffice işlemleri makro güvenliği yüksek, ayrı geçici bir profil kullanır. Şifreli belgelerin korumasını önce Office içinde kaldırın. Geçici PDF'ler normal kapanışta temizlenir.

**Yazıcıya gönderildi**, fiziksel baskının tamamlandığı anlamına gelmez. Çevrimdışı yazıcı, kâğıt ve mürekkep durumunu Windows yazıcı kuyruğundan izleyin. Çift taraf ve renk seçenekleri yazıcı desteğine bağlıdır. Microsoft Print to PDF gibi sanal yazıcılar çıktı konumu sorabilir.

**Kuyruğu durdur**, sonraki belgeleri durdurur. Yazıcıya aktarımı başlamış mevcut iş tamamlanır; gönderilmiş işleri Windows yazıcı kuyruğundan iptal edin. Aynı listeyi tekrar yazdırmak belgeleri yeniden gönderir. Dönüşüm zaman aşımında Office iletişim kutusu açık kalmışsa kapatıp yeniden deneyin. Başarısız yazdırma otomatik tekrarlanmaz.

## Doğrulama

```powershell
python -m unittest discover -s tests -v
python setup_dependencies.py --fixtures
$env:PRINTER_INTEGRATION='1'
python -m unittest discover -s tests -v
```

Entegrasyon testleri modern Office örneklerini oluşturur, Apache POI'nin eski biçim örnekleriyle birlikte altı biçimi PDF’ye dönüştürür, PDF içeriğini ve kaynakların değişmediğini kontrol eder, yazıcıları listeler ve uygulamada önizlemeyi açar. Fiziksel baskı yapmaz. Testler için python-docx, openpyxl, python-pptx, pypdf ve Pillow gerekir; bu bilgisayarın test ortamında bulunurlar. Kanıtlar `test-output` klasörüne kaydedilir.

### Bu bilgisayardaki doğrulama sınırı

Belge dönüşümü ve önizleme doğrulandı. Geliştirme oturumundan gerçek yazıcı açma çağrısı Windows tarafından `ERROR_ACCESS_DENIED (5)` ile reddedildi; sanal PDF yazıcısına da iş gönderilemedi. Bu nedenle uçtan uca baskı doğrulanmış değildir. Uygulama bu hatada iş gönderildi demek yerine açık bir izin hatası gösterir. Windows/kurum yazıcı izinleri uygun bir oturumda baskı testi yapılmalıdır. Güvenlik politikaları otomatik değiştirilmez.

PDF görüntüleme/yazdırma: [SumatraPDF 3.6.1](https://www.sumatrapdfreader.org/), dijital imzalı resmi taşınabilir dağıtım. Lisans/kaynak: [SumatraPDF deposu](https://github.com/sumatrapdfreader/sumatrapdf). SumatraPDF uygulamadan ayrı bir işlem olarak çalışır.
