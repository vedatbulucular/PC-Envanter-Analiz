# PC Envanter Analiz Sistemi 💻📊

<p align="center">
  <img src="arayuz_tasarimi.png" alt="PC Envanter Analiz Sistemi Arayüzü">
</p>

PC Envanter Analiz Sistemi; çalıştırıldığı Windows bilgisayarın donanım, yazılım, lisans, ağ ve aygıt durumunu tek işlemle analiz eden, sonuçları kullanıcı arayüzünde gösteren ve çok sekmeli Excel raporu oluşturan taşınabilir bir masaüstü uygulamasıdır. Proje, Çayırova Belediyesi Bilgi İşlem Müdürlüğündeki bilgisayar inceleme ve envanter kontrol süreçlerini hızlandırmak amacıyla zorunlu yaz stajı kapsamında geliştirilmiştir.

## 🎯 Proje Amacı

Saha teknisyenlerinin bilgisayar adı, işlemci, RAM, disk kullanımı, IP yapılandırması, etki alanı üyeliği, lisans durumu, zorunlu yazılımlar ve sorunlu aygıtlar gibi bilgileri farklı Windows araçlarından manuel olarak kontrol etmesi zaman kaybına ve insan hatası riskine yol açabilmektedir.

Bu proje ile söz konusu kontrollerin tek bir uygulamada otomatikleştirilmesi, sonuçların standartlaştırılması ve teknik personelin kısa sürede incelenebilir bir Excel raporu elde etmesi amaçlanmıştır.

## ✨ Öne Çıkan Özellikler

- **Eşzamanlı ve çok iş parçacıklı analiz:** `concurrent.futures.ThreadPoolExecutor` kullanılarak 16 bağımsız sistem sorgusu iş parçacığı havuzunda eşzamanlı olarak yürütülür. Uzun süren sorgular ana kullanıcı arayüzü iş parçacığından ayrıldığı için arayüz analiz sırasında yanıt verebilir durumda kalır.
- **WMI ve COM iş parçacığı güvenliği:** WMI kullanan görevlerin her iş parçacığında güvenli biçimde çalışması için özel `@with_com` decorator yapısı geliştirilmiştir. Bu yapı `pythoncom.CoInitialize()` ve `pythoncom.CoUninitialize()` çağrılarını yönetir.
- **Donanım analizi:** Bilgisayar adı, işletim sistemi sürümü, işlemci modeli, adres genişliği (32/64-bit), fiziksel ve mantıksal çekirdek sayıları, RAM bilgileri ve disk doluluk oranları analiz edilir.
- **Güvenli disk taraması:** Erişim izni bulunmayan, korumalı veya sorgulanamayan disk birimlerinde oluşan istisnalar (`PermissionError` dâhil) birim bazında yakalanır; diğer disklerin analizi kesintiye uğramadan devam eder.
- **Ağ yapılandırması denetimi:** Fiziksel ağ bağdaştırıcıları WMI üzerinden incelenerek sistemin statik IP mi yoksa DHCP tabanlı otomatik IP mi kullandığı belirlenir. Statik IP tespitinde kullanıcı arayüzünde görsel uyarı oluşturulur.
- **Etki alanı üyeliği kontrolü:** Bilgisayarın bir etki alanına mı yoksa çalışma grubuna mı bağlı olduğu tespit edilir.
- **Yazılım kontrolü:** Windows Kayıt Defteri ve bilinen kurulum dizinleri üzerinden Google Chrome, Microsoft Office, TeamViewer, Bitdefender Endpoint Security Tools, Adobe Acrobat Reader ve sıkıştırma araçları denetlenir.
- **Windows ve Office lisans denetimi:** WMI `SoftwareLicensingProduct` sınıfı kullanılarak Windows ve Microsoft Office etkinleştirme durumları kontrol edilir.
- **Sorunlu aygıt tespiti:** WMI `Win32_PnPEntity` sınıfında `ConfigManagerErrorCode != 0` olan yapılandırma hatalı aygıtlar belirlenir; aygıt adı, PnP aygıt kimliği (`PNPDeviceID`) ve hata kodu raporlanır.
- **Akıllı değerlendirme sistemi:** RAM kapasitesi, disk doluluğu, lisans durumu, ağ üyeliği ve zorunlu yazılımlar belirlenen kurallara göre değerlendirilerek uyarı listesi oluşturulur.
- **Esnek kullanıcı arayüzü:** `customtkinter` ile geliştirilen karanlık temalı arayüz, farklı pencere genişliklerine uyum sağlayan grid yapısı ve uzun metinleri alt satıra aktaran dinamik `wraplength` yönetimi içerir.
- **Çok sekmeli Excel raporu:** `openpyxl` kullanılarak `pc_rapor_{bilgisayar_adi}.xlsx` adlı rapor oluşturulur. Çalışma kitabı `Sistem Ozeti`, `Disk Bilgileri`, `Program Durumu`, `Eksik Suruculer` ve `Degerlendirme` adlı beş çalışma sayfasından oluşur.

## 🛠️ Kullanılan Teknolojiler

- **Programlama dili:** Python 3
- **Kullanıcı arayüzü:** `customtkinter`, `tkinter`
- **Sistem ve donanım sorguları:** `wmi`, `psutil`, `platform`, `socket`
- **Windows Kayıt Defteri:** `winreg`
- **Eşzamanlılık ve COM yönetimi:** `threading`, `concurrent.futures`, `pythoncom`
- **Excel raporlama:** `openpyxl`
- **Paketleme ve dağıtım:** `PyInstaller`, PowerShell
- **Sürüm kontrolü:** Git ve GitHub

## 📁 Proje Yapısı

```text
PC_Envanter_Analiz/
├── arayuz.py                  # Kullanıcı arayüzü ve analiz koordinasyonu
├── sistem_bilgisi.py          # Sistem sorguları, değerlendirme ve Excel raporu
├── build.ps1                  # PyInstaller üretim derleme betiği
├── arayuz_tasarimi.png        # README arayüz görseli
├── README.md                  # Proje dokümantasyonu
├── .gitignore                 # Sürüm kontrolü dışında tutulacak dosyalar
└── dist/
    └── PC_Envanter_Analiz.exe # Nihai çalıştırılabilir uygulama
```

## 🚀 Kullanım

Uygulama, hedef bilgisayarda Python kurulumu gerektirmeden çalışabilmesi için PyInstaller ile `--onefile`, `--windowed` ve `--uac-admin` parametreleri kullanılarak yaklaşık 20,4 MB boyutunda tek bir `.exe` dosyası olarak paketlenmiştir.

1. `dist` klasöründeki `PC_Envanter_Analiz.exe` dosyasını çalıştırın.
2. Windows tarafından gösterilen yönetici izni penceresini onaylayın.
3. Arayüz açıldığında **Sistemi Analiz Et** butonuna tıklayın.
4. Özet kartlarından ve analiz çıktısı alanından sonuçları inceleyin.
5. Sonuçları kaydetmek için **Excel'e Aktar** butonuna tıklayın.
6. Oluşturulan `pc_rapor_{bilgisayar_adi}.xlsx` dosyasını uygulamanın çalıştırıldığı dizinde görüntüleyin.

## ℹ️ Kapsam ve Notlar

- Uygulama uzaktaki bilgisayarları ağ üzerinden taramaz; çalıştırıldığı Windows bilgisayarı yerel olarak analiz eder.
- Statik IP kullanımı tek başına bir arıza değildir. Kurum standartlarına göre kontrol edilmesi gereken bir durum olarak uyarı listesine eklenir.
- `ConfigManagerErrorCode != 0` koşulu, yalnızca sürücüsü bulunmayan aygıtları değil, farklı yapılandırma sorunlarına sahip aygıtları da kapsayabilir.
- Oluşturulan rapor; bilgisayar adı, ağ üyeliği ve aygıt kimliği gibi teknik bilgiler içerebilir. Raporlar kurumun bilgi güvenliği kurallarına uygun biçimde saklanmalıdır.

---

**Geliştirici:** Vedat Bulucular  
**Kurum:** Çayırova Belediyesi Bilgi İşlem Müdürlüğü  
**Üniversite:** Düzce Üniversitesi – Bilgisayar Mühendisliği
