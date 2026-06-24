# PC Envanter Analiz Sistemi 💻📊

Bu proje, kurum içi bilgisayarların donanım, yazılım ve ağ özelliklerini tek bir tıkla analiz edebilen, raporlayan ve Excel formatında dışa aktarabilen modern bir masaüstü uygulamasıdır. Staj projesi kapsamında geliştirilmiştir.

## 🎯 Proje Amacı
Ağdaki bilgisayarların özelliklerini manuel olarak kontrol etmenin getirdiği zaman kaybını ve insan hatası riskini ortadan kaldırmak; cihazların donanım (RAM, Disk, CPU) yükseltme ihtiyaçlarını ve kritik yazılımların (Antivirüs, Office) varlığını hızlıca tespit edip raporlamaktır.

## ✨ Özellikler
* **Detaylı Donanım Taraması:** Bilgisayar adı, OS sürümü, CPU çekirdek ve kullanım bilgileri ile birlikte RAM ve Disk doluluk oranlarını otomatik çeker.
* **Yazılım & Güvenlik Kontrolü:** Windows Kayıt Defteri (Registry) ve WMI (Windows Management Instrumentation) kullanarak Google Chrome, Microsoft Office ve Antivirüs (Windows Defender vb.) yazılımlarının kurulu olup olmadığını tespit eder.
* **Ağ (IP) Yapılandırma Tespiti:** WMI (Get-NetIPConfiguration mantığı) ile ağ bağdaştırıcısının IP atama yöntemini (Otomatik/DHCP veya Statik IP) bulur.
* **Akıllı Değerlendirme & Uyarı Sistemi:**
  * RAM 8 GB altındaysa *[!] RAM Yükseltme Gerekli* uyarısı.
  * Office yazılımı yoksa *[!] Eksik Yazılım* uyarısı.
  * Disk doluluk oranı %85'in üzerindeyse *[!] Disk Dikkat* uyarısı.
* **Dinamik Raporlama:** Tüm bu verileri `pandas` ve `openpyxl` kullanarak otomatik olarak `pc_rapor_{bilgisayar_adi}.xlsx` isminde yapılandırılmış bir Excel dosyasına aktarır.
* **Modern Arayüz:** `customtkinter` ile geliştirilmiş, kullanıcı dostu, Dark Mode (karanlık tema) destekli, durum bildirimli şık bir arayüze sahiptir.

## 🛠️ Kullanılan Teknolojiler
* **Dil:** Python 3
* **Arayüz (GUI):** `customtkinter`, `tkinter`
* **Arka Uç & Sistem Sorguları:** `psutil`, `wmi`, `winreg`, `socket`, `platform`
* **Veri İşleme ve Dışa Aktarım:** `pandas`, `openpyxl`
* **Dağıtım (Build):** `PyInstaller` (Bağımsız, kurulum gerektirmeyen .exe oluşturmak için)
* **Senkronizasyon:** `threading` ve `pythoncom` (Arayüzün donmasını engellemek için)

## 🚀 Kullanım Şekli
Uygulama, sistemler üzerinde kurulum gerektirmeden çalışabilmesi için tek bir klasör (veya tek dosya) `.exe` mimarisinde derlenmiştir.

1. Projeyi indirin veya kopyalayın.
2. `dist` (veya `dist/PC_Envanter_Analiz`) klasörü içerisindeki `PC_Envanter_Analiz.exe` dosyasını çalıştırın. *(Uygulama WMI sorguları yapacağından Antivirüs tarafından engellenmemesi adına Yönetici İzinleri gerekebilir).*
3. Arayüz açıldığında **"Sistemi Analiz Et"** butonuna tıklayın.
4. Çıkan sonuçları ekranda görüntüleyin. IP durumu (Statik/DHCP) veya disk uyarıları otomatik renklendirilmiş kartlarda görülecektir.
5. Verileri saklamak için **"Excel'e Aktar"** butonuna tıklayın; uygulama bulunduğunuz dizine raporu kaydedecektir.

---

