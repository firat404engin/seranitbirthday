using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using ExcelDataReader;
using System.Text;  // Encoding için
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing;
using System.Collections.Generic; // List için
using System.Linq; // FirstOrDefault için

namespace DogumGunuHatirlatici
{
    public partial class Form1 : Form
    {
        // Outlook hesaplarını saklamak için liste
        private List<OutlookAccount> outlookAccounts = new List<OutlookAccount>();
        private Outlook.Application outlookApp;
        
        // Özel domain-e-posta eşleştirmeleri
        private Dictionary<string, string> domainToEmailMap = new Dictionary<string, string>
        {
            { "seranit.com.tr", "ik@seranit.com.tr" },
            { "vanucci.com", "insan.kaynaklari@vanucci.com" },
            { "mikrons.com.tr", "insan.kaynaklari@mikrons.com.tr" }
        };

        // Outlook hesabı için yardımcı sınıf
        private class OutlookAccount
        {
            public string DisplayName { get; set; }
            public string EmailAddress { get; set; }
            public Outlook.Account Account { get; set; }

            public override string ToString()
            {
                return DisplayName;
            }
        }

        public Form1()
        {
            InitializeComponent();

            // Encoding provider sadece bir kere kayıt edilmeli
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            
            // Modern görünüm için kontrolleri hazırla
            ApplyModernDesign();
            
            // Form kapatıldığında Outlook uygulamasını kapat
            this.FormClosing += Form1_FormClosing;
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            // Outlook hesaplarını yükle
            LoadOutlookAccounts();
            
            // Desteklenen domainler hakkında bilgi ver
            UpdateSupportedDomainsInfo();
        }

        private void LoadOutlookAccounts()
        {
            try
            {
                Log("Outlook hesapları yükleniyor...");
                
                // Hesap listesini temizle
                outlookAccounts.Clear();
                
                // Mevcut Outlook uygulaması varsa serbest bırak
                if (outlookApp != null)
                {
                    try
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
                    }
                    catch { /* Hata olursa sessizce devam et */ }
                    outlookApp = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                
                // Outlook uygulamasını başlat
                outlookApp = new Outlook.Application();
                
                // Outlook hesaplarını al
                Outlook.Accounts accounts = outlookApp.Session.Accounts;
                
                if (accounts.Count == 0)
                {
                    Log("UYARI: Outlook'ta yapılandırılmış hesap bulunamadı.");
                    MessageBox.Show("Outlook'ta yapılandırılmış hesap bulunamadı. Lütfen Outlook'ta en az bir hesap yapılandırın.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                // Her hesabı listeye ekle
                foreach (Outlook.Account account in accounts)
                {
                    var outlookAccount = new OutlookAccount
                    {
                        DisplayName = account.DisplayName,
                        EmailAddress = account.SmtpAddress,
                        Account = account
                    };
                    
                    outlookAccounts.Add(outlookAccount);
                }
                
                // Desteklenen domain'lere karşılık gelen hesaplar mevcut mu kontrol et
                foreach (var domainPair in domainToEmailMap)
                {
                    var domain = domainPair.Key;
                    var email = domainPair.Value;
                    
                    var accountExists = outlookAccounts.Any(a => 
                        string.Equals(a.EmailAddress, email, StringComparison.OrdinalIgnoreCase));
                    
                    if (!accountExists)
                    {
                        Log($"UYARI: @{domain} için kullanılacak {email} hesabı Outlook'ta bulunamadı!");
                    }
                    else
                    {
                        Log($"@{domain} için {email} hesabı kullanıma hazır");
                    }
                }
                
                Log($"Toplam {outlookAccounts.Count} Outlook hesabı yüklendi.");
                
                // UI bilgisini güncelle
                UpdateAccountInfo();
            }
            catch (Exception ex)
            {
                Log($"HATA: Outlook hesapları yüklenirken bir hata oluştu: {ex.Message}");
                MessageBox.Show($"Outlook hesapları yüklenirken bir hata oluştu: {ex.Message}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void UpdateAccountInfo()
        {
            // Kullanılabilir hesapları listele
            if (outlookAccounts.Count > 0)
            {
                mail.Text = string.Join(", ", outlookAccounts.Select(a => a.EmailAddress));
            }
            else
            {
                mail.Text = "Outlook hesabı bulunamadı";
            }
        }
        
        private void UpdateSupportedDomainsInfo()
        {
            Log("Doğum günü e-postaları, alıcının domainine göre otomatik olarak şu hesaplardan gönderilecek:");
            foreach (var domain in domainToEmailMap.Keys)
            {
                Log($"  * @{domain} alıcıları için -> {domainToEmailMap[domain]} hesabı kullanılacak");
            }
            Log("Desteklenmeyen domainlere sahip alıcılara e-posta gönderilmeyecek!");
        }

        private string GetDomainFromEmail(string email)
        {
            if (string.IsNullOrEmpty(email))
                return string.Empty;

            // E-posta adresinden domain kısmını çıkar (@'den sonraki kısım)
            int atIndex = email.IndexOf('@');
            if (atIndex >= 0 && atIndex < email.Length - 1)
                return email.Substring(atIndex + 1).ToLower();

            return string.Empty;
        }

        private void btnRefreshOutlook_Click(object sender, EventArgs e)
        {
            LoadOutlookAccounts();
        }

        private void ApplyModernDesign()
        {
            // Form kenarlıklarını yumuşat
            this.FormBorderStyle = FormBorderStyle.Sizable;
            
            // Form gölgesi ekle
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            
            // Butonlara hover efekti ekle
            foreach (Control control in this.Controls)
            {
                if (control is Button btn)
                {
                    btn.FlatAppearance.MouseOverBackColor = ControlPaint.Light(btn.BackColor);
                    btn.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(btn.BackColor);
                    btn.FlatAppearance.BorderColor = btn.BackColor;
                    btn.FlatAppearance.BorderSize = 0;
                }
                
                // Panel içindeki butonlar için de aynı efekti uygula
                if (control is Panel panel)
                {
                    ApplyEffectsToPanel(panel);
                }
            }
            
            // Ana panel için de efektleri uygula
            ApplyEffectsToPanel(panelMain);
            
            // ListView için özel stil
            LstOnizleme.GridLines = false;
            LstOnizleme.BorderStyle = BorderStyle.FixedSingle;
            LstOnizleme.FullRowSelect = true;
            
            // Form başlık rengini değiştir
            this.Text = "Doğum Günü Hatırlatıcı - Modern Arayüz";
        }
        
        private void ApplyEffectsToPanel(Panel panel)
        {
            foreach (Control control in panel.Controls)
            {
                if (control is Button btn)
                {
                    btn.FlatAppearance.MouseOverBackColor = ControlPaint.Light(btn.BackColor);
                    btn.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(btn.BackColor);
                    btn.FlatAppearance.BorderColor = btn.BackColor;
                    btn.FlatAppearance.BorderSize = 0;
                    
                    // Butonlara gölge efekti ekle
                    btn.Paint += (sender, e) => 
                    {
                        Button b = sender as Button;
                        ControlPaint.DrawBorder(e.Graphics, b.ClientRectangle, 
                            Color.LightGray, 0, ButtonBorderStyle.Solid,
                            Color.LightGray, 0, ButtonBorderStyle.Solid,
                            Color.DarkGray, 1, ButtonBorderStyle.Solid,
                            Color.DarkGray, 1, ButtonBorderStyle.Solid);
                    };
                }
                
                if (control is TextBox txt)
                {
                    txt.BorderStyle = BorderStyle.FixedSingle;
                }
            }
        }

        private void btnSelectExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDlg = new OpenFileDialog();
            openFileDlg.Filter = "Excel Files|*.xlsx;*.xls";
            openFileDlg.Title = "Doğum Günü Listesi Excel Dosyasını Seçiniz";

            if (openFileDlg.ShowDialog() == DialogResult.OK)
            {
                txtExcelPath.Text = openFileDlg.FileName;
                Log("Excel dosyası seçildi: " + openFileDlg.FileName);
            }
            openFileDlg.Dispose();
        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.Description = "Firma Görsellerinin Bulunduğu Klasörü Seçiniz";

            if (folderDlg.ShowDialog() == DialogResult.OK)
            {
                txtImageFolder.Text = folderDlg.SelectedPath;
                Log("Görsel klasörü seçildi: " + folderDlg.SelectedPath);
            }
            folderDlg.Dispose();
        }


        private void Log(string message)
        {
            string timeStamp = $"[{DateTime.Now:HH:mm:ss}]";
            string logMessage = $"{timeStamp} {message}";
            
            lstLog.Items.Add(logMessage);
            lstLog.TopIndex = lstLog.Items.Count - 1; // Otomatik scroll
        }

        private DataTable ReadExcel(string filePath)
        {
            // Encoding provider zaten constructor'da kayıt edildi, burada tekrar yazmaya gerek yok.

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var conf = new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true }
                    };

                    var result = reader.AsDataSet(conf);
                    return result.Tables[0];
                }
            }
        }

        private bool IsBirthdayTomorrow(DateTime birthDate)
        {
            DateTime tomorrow = DateTime.Today.AddDays(1);
            return (birthDate.Month == tomorrow.Month && birthDate.Day == tomorrow.Day);
        }

        private bool IsBirthdayInNextWeek(DateTime birthDate)
        {
            // Bugünden sonraki 7 gün içinde doğum günü olup olmadığını kontrol eder
            // Bugün doğum günü olanlar dahil edilmez
            for (int i = 1; i <= 7; i++)
            {
                DateTime checkDate = DateTime.Today.AddDays(i);
                if (birthDate.Month == checkDate.Month && birthDate.Day == checkDate.Day)
                {
                    return true;
                }
            }
            return false;
        }

        private bool IsStartDateOlderThan60Days(DateTime startDate)
        {
            return (DateTime.Today - startDate).TotalDays > 60;
        }

        private bool IsWeekend(DateTime date)
        {
            return date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday;
        }

        private bool IsBirthdayOnWeekend(DateTime birthDate)
        {
            // Doğum gününün bu yıl içinde hangi güne denk geldiğini kontrol eder
            // Gelecek 7 gün içindeki doğum günleri için
            for (int i = 1; i <= 7; i++)
            {
                DateTime checkDate = DateTime.Today.AddDays(i);
                if (birthDate.Month == checkDate.Month && birthDate.Day == checkDate.Day)
                {
                    // Doğum günü bu tarihe denk geliyor, hafta sonu mu kontrol et
                    return IsWeekend(checkDate);
                }
            }
            return false;
        }

        private string GetImagePath(string folderPath, string firma, bool isThImage)
        {
            string baseName = firma.ToLower();
            if (isThImage)
                baseName += "h";

            string pngPath = Path.Combine(folderPath, baseName + ".png");
            if (File.Exists(pngPath))
                return pngPath;

            string jpgPath = Path.Combine(folderPath, baseName + ".jpg");
            if (File.Exists(jpgPath))
                return jpgPath;

            return null;
        }

        private string ResizeImageToFixedSize(string originalImagePath, int width, int height)
        {
            string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + Path.GetExtension(originalImagePath));

            using (var originalImage = Image.FromFile(originalImagePath))
            using (var resizedBitmap = new Bitmap(width, height))
            using (var graphics = Graphics.FromImage(resizedBitmap))
            {
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;

                graphics.DrawImage(originalImage, 0, 0, width, height);

                resizedBitmap.Save(tempFile, ImageFormat.Png);
            }

            return tempFile;
        }

        private bool SendBirthdayMail(string toMail, string ccMail, string subject, string imagePath, DateTime dogumTarihi)
        {
            try
            {
                // Her e-posta için domaine uygun hesabı otomatik seç
                string domain = GetDomainFromEmail(toMail);
                string senderEmail = null;
                Outlook.Account selectedAccount = null;
                
                // Domaine göre uygun hesabı bul
                if (!string.IsNullOrEmpty(domain) && domainToEmailMap.ContainsKey(domain))
                {
                    senderEmail = domainToEmailMap[domain];
                    selectedAccount = outlookAccounts.FirstOrDefault(a => 
                        string.Equals(a.EmailAddress, senderEmail, StringComparison.OrdinalIgnoreCase))?.Account;
                }
                
                // Eğer domain eşleşmesi yoksa veya hesap bulunamadıysa, uyarı ver ve gönderme
                if (selectedAccount == null)
                {
                    string errorMsg = !string.IsNullOrEmpty(senderEmail) 
                        ? $"HATA: Domain '{domain}' için '{senderEmail}' hesabı Outlook'ta bulunamadı." 
                        : $"HATA: '{toMail}' için uygun bir gönderici hesabı bulunamadı.";
                    
                    Log(errorMsg);
                    return false;
                }

                // Outlook uygulaması zaten başlatıldı, tekrar başlatmaya gerek yok
                var mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                // Seçilen hesabı kullan
                mailItem.SendUsingAccount = selectedAccount;
                Log($"Mail, {selectedAccount.SmtpAddress} hesabı üzerinden gönderilecek - Alıcı: {toMail}");

                mailItem.To = toMail;
                
                // CC için noktalı virgülle ayrılmış e-posta adreslerini aynen kullan
                if (!string.IsNullOrWhiteSpace(ccMail))
                {
                    mailItem.CC = ccMail;
                }
                
                mailItem.Subject = subject;

                string contentId = "birthdayImage";

                string resizedImagePath = null;
                if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
                {
                    resizedImagePath = ResizeImageToFixedSize(imagePath, 540, 675);
                }

                // Sadece görsel içeren e-posta şablonu
                string htmlBody = $@"
                <html>
                    <head>
                        <style>
                            body {{
                                margin: 0;
                                padding: 0;
                                background-color: #ffffff;
                            }}
                            .image-container {{
                                text-align: center;
                                margin: 0 auto;
                                padding: 0;
                            }}
                        </style>
                    </head>
                    <body>
                        <div class='image-container'>
                            {(resizedImagePath != null ? $"<img src=\"cid:{contentId}\" style='max-width:100%;' />" : "")}
                        </div>
                    </body>
                </html>";

                mailItem.HTMLBody = htmlBody;

                if (resizedImagePath != null && File.Exists(resizedImagePath))
                {
                    Outlook.Attachment inlineAttachment = mailItem.Attachments.Add(resizedImagePath,
                        Outlook.OlAttachmentType.olByValue,
                        mailItem.Body.Length + 1,
                        "Birthday Image");

                    inlineAttachment.PropertyAccessor.SetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                        contentId
                    );
                }

                // Mail gönderimi sırasında oluşabilecek hataları yakalamak için try-catch bloğu
                try
                {
                    mailItem.Send();
                    
                    // Geçici dosyayı sil
                    if (resizedImagePath != null && File.Exists(resizedImagePath))
                        File.Delete(resizedImagePath);
                    
                    // CC'deki e-posta adreslerinin sayısını belirt
                    string ccInfo = string.IsNullOrWhiteSpace(ccMail) ? "" : $" (CC: {ccMail.Split(';').Length} kişi)";
                    Log($"Mail başarıyla gönderildi: {toMail}{ccInfo} - Doğum Günü: {dogumTarihi.ToString("dd MMMM", new System.Globalization.CultureInfo("tr-TR"))}");
                    return true;
                }
                catch (Exception ex)
                {
                    Log($"HATA: Mail gönderimi başarısız - Alıcı: {toMail} - Hata: {ex.Message}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log($"HATA: Mail oluşturma hatası - Alıcı: {toMail} - Hata: {ex.Message}");
                return false;
            }
        }

        private bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        private bool AreValidEmails(string emailList)
        {
            if (string.IsNullOrWhiteSpace(emailList))
                return true;

            string[] emails = emailList.Split(';');
            foreach (string email in emails)
            {
                if (!IsValidEmail(email.Trim()))
                    return false;
            }
            return true;
        }

        private void btnOnizleme_Click(object sender, EventArgs e)
        {
            string excelFile = txtExcelPath.Text;
            string imagesFolder = txtImageFolder.Text;

            if (!File.Exists(excelFile))
            {
                MessageBox.Show("Lütfen geçerli bir Excel dosyası seçiniz.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Directory.Exists(imagesFolder))
            {
                MessageBox.Show("Lütfen geçerli bir görsel klasörü seçiniz.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            LstOnizleme.Items.Clear();
            Log("Önizleme başlatıldı...");

            DataTable dt = ReadExcel(excelFile);
            int dogumGunuSayisi = 0;
            int hataliMailSayisi = 0;
            
            // Domain başına gruplanmış e-posta adresleri
            Dictionary<string, int> domainCounts = new Dictionary<string, int>();
            // Hangi hesaplardan gönderim yapılacağını takip et
            Dictionary<string, int> accountsToUse = new Dictionary<string, int>();

            foreach (DataRow row in dt.Rows)
            {
                string firma = row["firma"].ToString().Trim();
                string mail = row["mail"].ToString().Trim();
                string adsoyad = row["adsoyad"].ToString().Trim();
                string mudurMail = row["mudur"].ToString().Trim();

                // Domain istatistiği topla
                string domain = GetDomainFromEmail(mail);
                if (!string.IsNullOrEmpty(domain))
                {
                    if (domainCounts.ContainsKey(domain))
                        domainCounts[domain]++;
                    else
                        domainCounts[domain] = 1;
                    
                    // Kullanılacak hesabı belirle
                    string accountToUse = domainToEmailMap.ContainsKey(domain) ? domainToEmailMap[domain] : "diğer";
                    if (accountsToUse.ContainsKey(accountToUse))
                        accountsToUse[accountToUse]++;
                    else
                        accountsToUse[accountToUse] = 1;
                }

                // Mail formatı kontrolü
                bool mailGecerli = IsValidEmail(mail);
                bool mudurMailGecerli = AreValidEmails(mudurMail);

                if (!mailGecerli)
                {
                    Log($"HATA: Geçersiz e-posta adresi: {mail} (Kişi: {adsoyad})");
                    hataliMailSayisi++;
                }

                if (!mudurMailGecerli)
                {
                    Log($"HATA: Geçersiz müdür e-posta adresi: {mudurMail} (Kişi: {adsoyad})");
                    hataliMailSayisi++;
                }

                if (!DateTime.TryParse(row["dogumtarihi"].ToString(), out DateTime dogumTarihi))
                {
                    Log($"HATA: Geçersiz doğum tarihi: {adsoyad}");
                    continue;
                }

                if (!DateTime.TryParse(row["baslamatarihi"].ToString(), out DateTime baslamaTarihi))
                {
                    Log($"HATA: Geçersiz başlama tarihi: {adsoyad}");
                    continue;
                }

                string imagePath = null;
                bool baslama60GunUstu = IsStartDateOlderThan60Days(baslamaTarihi);
                bool dogumGunuVar = false;
                string dogumGunuTarihi = "";
                bool dogumGunuHaftaSonu = false;

                if (IsBirthdayInNextWeek(dogumTarihi))
                {
                    dogumGunuVar = true;
                    dogumGunuSayisi++;
                    // Doğum gününün hangi tarihte olduğunu bul
                    for (int i = 1; i <= 7; i++)
                    {
                        DateTime checkDate = DateTime.Today.AddDays(i);
                        if (dogumTarihi.Month == checkDate.Month && dogumTarihi.Day == checkDate.Day)
                        {
                            dogumGunuTarihi = checkDate.ToString("dd.MM.yyyy (dddd)", new System.Globalization.CultureInfo("tr-TR"));
                            dogumGunuHaftaSonu = IsWeekend(checkDate);
                            break;
                        }
                    }

                    // Doğum günü hafta sonuna denk geliyorsa veya başlama tarihi 60 günden az ise "h" ekli resim kullan
                    if (dogumGunuHaftaSonu || !baslama60GunUstu)
                    {
                        imagePath = GetImagePath(imagesFolder, firma, true); // "h" ekli resim
                    }
                    else
                    {
                        // Doğum günü hafta içi ve başlama tarihi 60 günden fazla ise normal resim
                        imagePath = GetImagePath(imagesFolder, firma, false);
                        // Normal resim yoksa "h" ekli resmi dene
                        if (imagePath == null)
                            imagePath = GetImagePath(imagesFolder, firma, true);
                    }
                }

                // Sadece doğum günü gelecek 7 gün içinde olanları listele
                if (dogumGunuVar)
                {
                    // Göndericinin e-posta adresini belirle
                    string gondericiHesap = "Gönderilmeyecek";
                    if (domainToEmailMap.ContainsKey(domain))
                    {
                        gondericiHesap = domainToEmailMap[domain];
                        
                        // Eğer bu hesap Outlook'ta yoksa uyarı ver
                        bool hesapMevcut = outlookAccounts.Any(a => 
                            string.Equals(a.EmailAddress, gondericiHesap, StringComparison.OrdinalIgnoreCase));
                        
                        if (!hesapMevcut)
                        {
                            gondericiHesap += " (BULUNAMADI!)";
                        }
                    }

                    ListViewItem item = new ListViewItem(firma);
                    item.SubItems.Add(mail);
                    item.SubItems.Add(mudurMail);
                    item.SubItems.Add(imagePath != null ? Path.GetFileName(imagePath) : "Yok");
                    item.SubItems.Add(dogumGunuTarihi);
                    item.SubItems.Add(gondericiHesap); // Gönderici hesap bilgisini ekliyoruz

                    // Hafta sonu doğum günleri için arka plan rengini değiştir
                    if (dogumGunuHaftaSonu)
                    {
                        item.BackColor = System.Drawing.Color.LightYellow;
                    }

                    // Hatalı mail adresi varsa kırmızı renk ile vurgula
                    if (!mailGecerli || !mudurMailGecerli)
                    {
                        item.ForeColor = System.Drawing.Color.Red;
                    }
                    
                    // Gönderici hesabı bulunamadıysa, satırı farklı renkte göster
                    if (gondericiHesap.Contains("BULUNAMADI"))
                    {
                        item.ForeColor = System.Drawing.Color.Red;
                    }

                    LstOnizleme.Items.Add(item);

                    string resimTipi = dogumGunuHaftaSonu ? "Hafta Sonu (h)" : (baslama60GunUstu ? "Normal" : "h ekli");
                    string logSatir = $"Firma: {firma}, Mail: {mail}, CC: {mudurMail}, Resim: {(imagePath != null ? Path.GetFileName(imagePath) : "Yok")} ({resimTipi}), Doğum Günü: {dogumGunuTarihi}, Gönderici: {gondericiHesap}";
                    Log(logSatir);
                }
            }

            // Domain bazında istatistikleri göster
            if (domainCounts.Count > 0)
            {
                Log("Domain bazında alıcı dağılımı:");
                foreach (var pair in domainCounts.OrderByDescending(x => x.Value))
                {
                    Log($"  {pair.Key}: {pair.Value} kişi");
                }
            }
            
            // Kullanılacak hesapları göster
            if (accountsToUse.Count > 0)
            {
                Log("Gönderim için kullanılacak hesaplar:");
                foreach (var pair in accountsToUse.OrderByDescending(x => x.Value))
                {
                    if (pair.Key != "diğer")
                    {
                        Log($"  {pair.Key}: {pair.Value} e-posta");
                    }
                    else
                    {
                        Log($"  Desteklenmeyen domain e-postaları: {pair.Value} kişi - Gönderim yapılamayacak");
                    }
                }
            }

            // Gönder butonunu aktif et (sadece doğum günü varsa VE hatalı mail yoksa)
            btnSendMails.Enabled = dogumGunuSayisi > 0 && hataliMailSayisi == 0;
            
            if (dogumGunuSayisi > 0)
            {
                Log($"Önizleme tamamlandı. {dogumGunuSayisi} kişinin doğum günü bulundu.");
                lblHeader.Text = $"Önizleme ({dogumGunuSayisi} kişi)";
                
                if (hataliMailSayisi > 0)
                {
                    Log($"UYARI: {hataliMailSayisi} adet hatalı e-posta adresi tespit edildi. Lütfen düzeltiniz!");
                    MessageBox.Show($"{hataliMailSayisi} adet hatalı e-posta adresi tespit edildi. Lütfen Excel dosyasındaki e-posta adreslerini kontrol ediniz.", 
                        "Hatalı E-posta Adresleri", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                Log("Önizleme tamamlandı. Gelecek 7 gün içinde doğum günü bulunamadı.");
                lblHeader.Text = "Önizleme (0 kişi)";
            }
        }

        private void btnSendMails_Click(object sender, EventArgs e)
        {
            string excelFile = txtExcelPath.Text;
            string imagesFolder = txtImageFolder.Text;

            if (!File.Exists(excelFile))
            {
                MessageBox.Show("Lütfen geçerli bir Excel dosyası seçiniz.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Directory.Exists(imagesFolder))
            {
                MessageBox.Show("Lütfen geçerli bir görsel klasörü seçiniz.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Önizleme yapılmadıysa gönderim yapma
            if (LstOnizleme.Items.Count == 0)
            {
                MessageBox.Show("Lütfen önce önizleme yapınız.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            lstLog.Items.Clear();
            Log("E-posta gönderimi başladı...");

            // Domain bazında gönderim istatistikleri için Dictionary
            Dictionary<string, int> sentByDomain = new Dictionary<string, int>();
            // Hangi hesaplardan kaç mail gönderildiğini takip et
            Dictionary<string, int> sentByAccount = new Dictionary<string, int>();

            DataTable dt = ReadExcel(excelFile);
            int gonderimSayisi = 0;
            int hataliMailSayisi = 0;
            int gonderimHataSayisi = 0;

            foreach (DataRow row in dt.Rows)
            {
                string firma = row["firma"].ToString().Trim();
                string toMail = row["mail"].ToString().Trim();
                string adsoyad = row["adsoyad"].ToString().Trim();
                string mudurMail = row["mudur"].ToString().Trim();

                // Mail formatı kontrolü
                bool mailGecerli = IsValidEmail(toMail);
                bool mudurMailGecerli = AreValidEmails(mudurMail);

                if (!mailGecerli)
                {
                    Log($"HATA: Geçersiz e-posta adresi: {toMail} (Kişi: {adsoyad}) - E-posta gönderilmedi");
                    hataliMailSayisi++;
                    continue; // Hatalı mail varsa bu kişiye mail gönderme
                }

                if (!mudurMailGecerli)
                {
                    // Müdür maili hatalıysa, müdür mailini boş olarak ayarla
                    Log($"UYARI: Geçersiz müdür e-posta adresi: {mudurMail} (Kişi: {adsoyad}) - CC olmadan gönderilecek");
                    mudurMail = "";
                }

                if (!DateTime.TryParse(row["dogumtarihi"].ToString(), out DateTime dogumTarihi))
                {
                    continue;
                }

                if (!DateTime.TryParse(row["baslamatarihi"].ToString(), out DateTime baslamaTarihi))
                {
                    continue;
                }

                if (IsBirthdayInNextWeek(dogumTarihi))
                {
                    bool dogumGunuHaftaSonu = IsBirthdayOnWeekend(dogumTarihi);
                    bool baslama60GunUstu = IsStartDateOlderThan60Days(baslamaTarihi);

                    string imagePath = null;

                    // Doğum günü hafta sonuna denk geliyorsa veya başlama tarihi 60 günden az ise "h" ekli resim kullan
                    if (dogumGunuHaftaSonu || !baslama60GunUstu)
                    {
                        imagePath = GetImagePath(imagesFolder, firma, true); // "h" ekli resim
                    }
                    else
                    {
                        // Doğum günü hafta içi ve başlama tarihi 60 günden fazla ise normal resim
                        imagePath = GetImagePath(imagesFolder, firma, false);
                        // Normal resim yoksa "h" ekli resmi dene
                        if (imagePath == null)
                            imagePath = GetImagePath(imagesFolder, firma, true);
                    }

                    string subject = $"Doğum Gününüz Kutlu Olsun!";
                    
                    bool mailGonderildi = SendBirthdayMail(toMail, mudurMail, subject, imagePath, dogumTarihi);
                    
                    if (mailGonderildi)
                    {
                        gonderimSayisi++;
                        
                        // Domaini istatistik için kaydet
                        string domain = GetDomainFromEmail(toMail);
                        if (!string.IsNullOrEmpty(domain))
                        {
                            if (sentByDomain.ContainsKey(domain))
                                sentByDomain[domain]++;
                            else
                                sentByDomain[domain] = 1;
                        }
                        
                        // Kullanılan hesabı da istatistik için kaydet
                        string accountEmail = domainToEmailMap.ContainsKey(domain) ? domainToEmailMap[domain] : "diğer";
                        if (sentByAccount.ContainsKey(accountEmail))
                            sentByAccount[accountEmail]++;
                        else
                            sentByAccount[accountEmail] = 1;
                    }
                    else
                    {
                        gonderimHataSayisi++;
                    }
                }
            }
            
            // Gönderim tamamlandıktan sonra istatistikleri göster
            if (sentByDomain.Count > 0)
            {
                Log("Domain bazında gönderim istatistikleri:");
                foreach (var pair in sentByDomain.OrderByDescending(x => x.Value))
                {
                    Log($"  {pair.Key}: {pair.Value} e-posta");
                }
            }
            
            // Hesap bazında istatistikleri göster
            if (sentByAccount.Count > 0)
            {
                Log("Hesap bazında gönderim istatistikleri:");
                foreach (var pair in sentByAccount.OrderByDescending(x => x.Value))
                {
                    Log($"  {pair.Key}: {pair.Value} e-posta");
                }
            }

            if (hataliMailSayisi > 0)
            {
                Log($"UYARI: {hataliMailSayisi} adet hatalı e-posta adresi nedeniyle bazı e-postalar gönderilemedi.");
            }

            if (gonderimHataSayisi > 0)
            {
                Log($"UYARI: {gonderimHataSayisi} adet e-posta gönderim hatası oluştu.");
            }

            Log($"İşlem tamamlandı. {gonderimSayisi} kişiye e-posta başarıyla gönderildi.");
            
            // Gönderim tamamlandıktan sonra gönder butonunu devre dışı bırak
            btnSendMails.Enabled = false;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Outlook uygulamasını kapat ve kaynakları serbest bırak
            if (outlookApp != null)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
                }
                catch (Exception ex)
                {
                    // Hata oluşursa sessizce devam et
                    System.Diagnostics.Debug.WriteLine($"Outlook uygulaması kapatılırken hata: {ex.Message}");
                }
                finally
                {
                    outlookApp = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        private void panelMain_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
