🎉 Doğum Günü Hatırlatıcı Uygulaması 🎂
🔍 Genel Bakış
Doğum Günü Hatırlatıcı, şirket çalışanlarının doğum günlerini takip eden ve yaklaşan doğum günleri için otomatik e-posta gönderen bir Windows Forms uygulamasıdır. Çalışan bilgilerini Excel dosyasından okur ve Outlook entegrasyonu ile e-postaları otomatik gönderir. 📧🎈

✨ Özellikler
📊 Excel Entegrasyonu: Çalışan bilgileri Excel dosyasından okunur

📬 Outlook Entegrasyonu: Doğum günü e-postaları Outlook üzerinden gönderilir

🔄 Otomatik Hesap Seçimi: Alıcının domain adresine göre uygun gönderici hesabı seçilir

🖼️ Özelleştirilmiş Görseller: Firma bazlı özel doğum günü görselleri kullanılabilir

👀 Önizleme: Gönderilecek e-postalar detaylı önizlenebilir

📜 Gelişmiş Loglama: Tüm işlemler detaylı loglanır

🛠️ Teknik Gereksinimler
.NET Framework 4.7.2 veya üzeri

Microsoft Office Outlook (yüklü ve yapılandırılmış)

ExcelDataReader kütüphanesi

System.Text.Encoding.CodePages kütüphanesi

🚀 Kurulum
Projeyi derleyin veya derlenmiş dosyaları indirin

Uygulamayı çalıştırın

Outlook’un yüklü ve en az bir e-posta hesabı ile yapılandırılmış olduğundan emin olun

📚 Kullanım Kılavuzu
1️⃣ Excel Dosyasını Hazırlama
Excel dosyasında aşağıdaki sütunlar olmalıdır:

🏢 firma: Çalışanın firma adı

📧 mail: Çalışanın e-posta adresi

👤 adsoyad: Çalışanın adı soyadı

👔 mudur: Müdürün e-posta adresi (CC için)

🎂 dogumtarihi: Doğum tarihi (tarih formatında)

📅 baslamatarihi: İşe başlama tarihi (tarih formatında)

2️⃣ Görsel Dosyaları Hazırlama
🖼️ Her firma için bir görsel dosyası oluşturun

📁 Dosya adı firma adı ile aynı olmalı (ör. seranit.jpg)

📅 Hafta sonu doğum günleri için dosya adının sonuna h ekleyin (ör. seranith.jpg)

3️⃣ Uygulama Kullanımı
📂 Excel Dosyası Seç: "Dosya Seç" butonuyla dosyayı seçin

📁 Görsel Klasörü Seç: "Klasör Seç" butonuyla görsellerin olduğu klasörü seçin

👓 Önizleme: "Önizleme" butonuyla e-postaları kontrol edin

📤 Gönder: "GÖNDER" butonuyla e-postaları gönderin

4️⃣ Otomatik Hesap Seçimi
Domain bazlı gönderici hesapları:

✉️ @seranit.com.tr → seranit.com.tr

✉️ @vanucci.com → anucci.com

✉️ @mikrons.com.tr → mikrons.com.tr

⚠️ Desteklenmeyen domainlere e-posta gönderilmez.

🏗️ Kod Yapısı
📁 Ana Bileşenler
Form1.cs: Ana form ve iş mantığı

Form1.Designer.cs: Form tasarımı

ExcelDataReader: Excel okuma kütüphanesi

Microsoft.Office.Interop.Outlook: Outlook entegrasyonu

🔑 Önemli Metotlar
LoadOutlookAccounts(): Outlook hesaplarını yükler

ReadExcel(): Excel dosyasını okur

btnOnizleme_Click(): Önizleme işlemi

SendBirthdayMail(): E-posta gönderimi

IsBirthdayInNextWeek(): Doğum günü kontrolü

GetDomainFromEmail(): Email domain çıkarma

🎨 Özelleştirme
🔧 Domain-Hesap Eşleştirme: domainToEmailMap sözlüğünü Form1.cs içinde düzenleyin

✉️ E-posta Şablonu: SendBirthdayMail() metodundaki HTML şablonunu güncelleyin

🛠️ Hata Giderme
❌ Outlook Hesapları Yüklenmiyor:

Outlook yüklü ve yapılandırılmış mı kontrol edin

Yönetici olarak çalıştırmayı deneyin

❌ Excel Dosyası Okunamıyor:

Dosya formatı ve gerekli sütunlar doğru mu?

Dosya başka bir programda açık mı?

❌ E-postalar Gönderilmiyor:

Outlook hesaplarının yapılandırılması doğru mu?

Desteklenen domainlerden gönderiliyor mu?

🔐 Güvenlik Notları
Uygulama Outlook güvenlik modelini kullanır

Şifre ve hassas verileri Excel dosyasında tutmayın

Uygulamayı güvenilir kaynaklardan kullanın

🌟 Gelecek Geliştirmeler
💾 Veritabanı entegrasyonu

⏰ Otomatik zamanlama

📧 Ek e-posta şablonları

🌐 Çoklu dil desteği

💻 Web tabanlı arayüz

📜 Lisans
MIT Lisansı ile açık kaynak olarak dağıtılmaktadır.

📞 İletişim
Herhangi bir soru, öneri veya hata bildirimi için lütfen GitHub üzerinden iletişime geçin.

Bu dokümantasyon, Doğum Günü Hatırlatıcı uygulamasını kullanmanız ve geliştirmeniz için rehber niteliğindedir. 🚀🎉

