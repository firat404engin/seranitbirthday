namespace DogumGunuHatirlatici
{
    partial class Form1
    {
        /// <summary>
        ///Gerekli tasarımcı değişkeni.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///Kullanılan tüm kaynakları temizleyin.
        /// </summary>
        ///<param name="disposing">yönetilen kaynaklar dispose edilmeliyse doğru; aksi halde yanlış.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer üretilen kod

        /// <summary>
        /// Tasarımcı desteği için gerekli metot - bu metodun 
        ///içeriğini kod düzenleyici ile değiştirmeyin.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.lblExcel = new System.Windows.Forms.Label();
            this.txtExcelPath = new System.Windows.Forms.TextBox();
            this.btnSelectExcel = new System.Windows.Forms.Button();
            this.btnSelectFolder = new System.Windows.Forms.Button();
            this.txtImageFolder = new System.Windows.Forms.TextBox();
            this.lblImageFolder = new System.Windows.Forms.Label();
            this.btnSendMails = new System.Windows.Forms.Button();
            this.lstLog = new System.Windows.Forms.ListBox();
            this.LstOnizleme = new System.Windows.Forms.ListView();
            this.columnFirma = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnMail = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnMudur = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnResim = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnDogumGunu = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnGondericiHesap = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnOnizleme = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.lblHeader = new System.Windows.Forms.Label();
            this.panelTop = new System.Windows.Forms.Panel();
            this.lblAppTitle = new System.Windows.Forms.Label();
            this.panelMain = new System.Windows.Forms.Panel();
            this.btnRefreshOutlook = new System.Windows.Forms.Button();
            this.mail = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panelTop.SuspendLayout();
            this.panelMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblExcel
            // 
            this.lblExcel.AutoSize = true;
            this.lblExcel.Font = new System.Drawing.Font("Segoe UI Semibold", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.lblExcel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblExcel.Location = new System.Drawing.Point(20, 25);
            this.lblExcel.Name = "lblExcel";
            this.lblExcel.Size = new System.Drawing.Size(105, 23);
            this.lblExcel.TabIndex = 0;
            this.lblExcel.Text = "Excel Dosya:";
            // 
            // txtExcelPath
            // 
            this.txtExcelPath.BackColor = System.Drawing.Color.White;
            this.txtExcelPath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtExcelPath.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.txtExcelPath.Location = new System.Drawing.Point(144, 22);
            this.txtExcelPath.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtExcelPath.Name = "txtExcelPath";
            this.txtExcelPath.ReadOnly = true;
            this.txtExcelPath.Size = new System.Drawing.Size(318, 30);
            this.txtExcelPath.TabIndex = 2;
            // 
            // btnSelectExcel
            // 
            this.btnSelectExcel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.btnSelectExcel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSelectExcel.FlatAppearance.BorderSize = 0;
            this.btnSelectExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSelectExcel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnSelectExcel.ForeColor = System.Drawing.Color.White;
            this.btnSelectExcel.Location = new System.Drawing.Point(483, 22);
            this.btnSelectExcel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnSelectExcel.Name = "btnSelectExcel";
            this.btnSelectExcel.Size = new System.Drawing.Size(124, 30);
            this.btnSelectExcel.TabIndex = 3;
            this.btnSelectExcel.Text = "Dosya Seç";
            this.btnSelectExcel.UseVisualStyleBackColor = false;
            this.btnSelectExcel.Click += new System.EventHandler(this.btnSelectExcel_Click);
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.btnSelectFolder.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSelectFolder.FlatAppearance.BorderSize = 0;
            this.btnSelectFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSelectFolder.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnSelectFolder.ForeColor = System.Drawing.Color.White;
            this.btnSelectFolder.Location = new System.Drawing.Point(483, 65);
            this.btnSelectFolder.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(124, 30);
            this.btnSelectFolder.TabIndex = 6;
            this.btnSelectFolder.Text = "Klasör Seç";
            this.btnSelectFolder.UseVisualStyleBackColor = false;
            this.btnSelectFolder.Click += new System.EventHandler(this.btnSelectFolder_Click);
            // 
            // txtImageFolder
            // 
            this.txtImageFolder.BackColor = System.Drawing.Color.White;
            this.txtImageFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtImageFolder.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.txtImageFolder.Location = new System.Drawing.Point(144, 65);
            this.txtImageFolder.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtImageFolder.Name = "txtImageFolder";
            this.txtImageFolder.ReadOnly = true;
            this.txtImageFolder.Size = new System.Drawing.Size(318, 30);
            this.txtImageFolder.TabIndex = 5;
            // 
            // lblImageFolder
            // 
            this.lblImageFolder.AutoSize = true;
            this.lblImageFolder.Font = new System.Drawing.Font("Segoe UI Semibold", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.lblImageFolder.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblImageFolder.Location = new System.Drawing.Point(20, 68);
            this.lblImageFolder.Name = "lblImageFolder";
            this.lblImageFolder.Size = new System.Drawing.Size(99, 23);
            this.lblImageFolder.TabIndex = 4;
            this.lblImageFolder.Text = "Görsel Yolu:";
            // 
            // btnSendMails
            // 
            this.btnSendMails.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSendMails.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(204)))), ((int)(((byte)(113)))));
            this.btnSendMails.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSendMails.Enabled = false;
            this.btnSendMails.FlatAppearance.BorderSize = 0;
            this.btnSendMails.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSendMails.Font = new System.Drawing.Font("Segoe UI", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnSendMails.ForeColor = System.Drawing.Color.White;
            this.btnSendMails.Location = new System.Drawing.Point(1089, 578);
            this.btnSendMails.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnSendMails.Name = "btnSendMails";
            this.btnSendMails.Size = new System.Drawing.Size(153, 50);
            this.btnSendMails.TabIndex = 7;
            this.btnSendMails.Text = "GÖNDER";
            this.btnSendMails.UseVisualStyleBackColor = false;
            this.btnSendMails.Click += new System.EventHandler(this.btnSendMails_Click);
            // 
            // lstLog
            // 
            this.lstLog.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lstLog.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(45)))), ((int)(((byte)(48)))));
            this.lstLog.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.lstLog.Font = new System.Drawing.Font("Consolas", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.lstLog.ForeColor = System.Drawing.Color.LightGray;
            this.lstLog.FormattingEnabled = true;
            this.lstLog.ItemHeight = 20;
            this.lstLog.Location = new System.Drawing.Point(0, 0);
            this.lstLog.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.lstLog.Name = "lstLog";
            this.lstLog.Size = new System.Drawing.Size(1254, 140);
            this.lstLog.TabIndex = 8;
            // 
            // LstOnizleme
            // 
            this.LstOnizleme.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.LstOnizleme.BackColor = System.Drawing.Color.White;
            this.LstOnizleme.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.LstOnizleme.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnFirma,
            this.columnMail,
            this.columnMudur,
            this.columnResim,
            this.columnDogumGunu,
            this.columnGondericiHesap});
            this.LstOnizleme.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.LstOnizleme.FullRowSelect = true;
            this.LstOnizleme.GridLines = true;
            this.LstOnizleme.HideSelection = false;
            this.LstOnizleme.Location = new System.Drawing.Point(12, 25);
            this.LstOnizleme.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.LstOnizleme.Name = "LstOnizleme";
            this.LstOnizleme.Size = new System.Drawing.Size(1067, 195);
            this.LstOnizleme.TabIndex = 9;
            this.LstOnizleme.UseCompatibleStateImageBehavior = false;
            this.LstOnizleme.View = System.Windows.Forms.View.Details;
            // 
            // columnFirma
            // 
            this.columnFirma.Text = "Firma";
            this.columnFirma.Width = 120;
            // 
            // columnMail
            // 
            this.columnMail.Text = "Mail";
            this.columnMail.Width = 180;
            // 
            // columnMudur
            // 
            this.columnMudur.Text = "Müdür Mail";
            this.columnMudur.Width = 180;
            // 
            // columnResim
            // 
            this.columnResim.Text = "Resim";
            this.columnResim.Width = 120;
            // 
            // columnDogumGunu
            // 
            this.columnDogumGunu.Text = "Doğum Günü";
            this.columnDogumGunu.Width = 180;
            // 
            // columnGondericiHesap
            // 
            this.columnGondericiHesap.Text = "Gönderici Hesap";
            this.columnGondericiHesap.Width = 250;
            // 
            // btnOnizleme
            // 
            this.btnOnizleme.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOnizleme.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(52)))), ((int)(((byte)(152)))), ((int)(((byte)(219)))));
            this.btnOnizleme.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnOnizleme.FlatAppearance.BorderSize = 0;
            this.btnOnizleme.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOnizleme.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnOnizleme.ForeColor = System.Drawing.Color.White;
            this.btnOnizleme.Location = new System.Drawing.Point(1089, 389);
            this.btnOnizleme.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnOnizleme.Name = "btnOnizleme";
            this.btnOnizleme.Size = new System.Drawing.Size(112, 39);
            this.btnOnizleme.TabIndex = 10;
            this.btnOnizleme.Text = "Önizleme";
            this.btnOnizleme.UseVisualStyleBackColor = false;
            this.btnOnizleme.Click += new System.EventHandler(this.btnOnizleme_Click);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(45)))), ((int)(((byte)(48)))));
            this.panel1.Controls.Add(this.lstLog);
            this.panel1.Location = new System.Drawing.Point(9, 194);
            this.panel1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1254, 127);
            this.panel1.TabIndex = 11;
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel2.BackColor = System.Drawing.Color.White;
            this.panel2.Controls.Add(this.LstOnizleme);
            this.panel2.Location = new System.Drawing.Point(0, 364);
            this.panel2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1082, 277);
            this.panel2.TabIndex = 12;
            // 
            // lblHeader
            // 
            this.lblHeader.AutoSize = true;
            this.lblHeader.Font = new System.Drawing.Font("Segoe UI", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.lblHeader.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.lblHeader.Location = new System.Drawing.Point(3, 323);
            this.lblHeader.Name = "lblHeader";
            this.lblHeader.Size = new System.Drawing.Size(115, 31);
            this.lblHeader.TabIndex = 13;
            this.lblHeader.Text = "Önizleme";
            // 
            // panelTop
            // 
            this.panelTop.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.panelTop.Controls.Add(this.lblAppTitle);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(0, 0);
            this.panelTop.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(1254, 50);
            this.panelTop.TabIndex = 14;
            // 
            // lblAppTitle
            // 
            this.lblAppTitle.AutoSize = true;
            this.lblAppTitle.Font = new System.Drawing.Font("Segoe UI", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.lblAppTitle.ForeColor = System.Drawing.Color.White;
            this.lblAppTitle.Location = new System.Drawing.Point(12, 9);
            this.lblAppTitle.Name = "lblAppTitle";
            this.lblAppTitle.Size = new System.Drawing.Size(258, 31);
            this.lblAppTitle.TabIndex = 0;
            this.lblAppTitle.Text = "Doğum Günü Bildirimi";
            // 
            // panelMain
            // 
            this.panelMain.BackColor = System.Drawing.Color.White;
            this.panelMain.Controls.Add(this.btnRefreshOutlook);
            this.panelMain.Controls.Add(this.mail);
            this.panelMain.Controls.Add(this.label1);
            this.panelMain.Controls.Add(this.btnSelectFolder);
            this.panelMain.Controls.Add(this.txtImageFolder);
            this.panelMain.Controls.Add(this.lblImageFolder);
            this.panelMain.Controls.Add(this.btnSelectExcel);
            this.panelMain.Controls.Add(this.txtExcelPath);
            this.panelMain.Controls.Add(this.lblExcel);
            this.panelMain.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelMain.Location = new System.Drawing.Point(0, 50);
            this.panelMain.Margin = new System.Windows.Forms.Padding(4);
            this.panelMain.Name = "panelMain";
            this.panelMain.Size = new System.Drawing.Size(1254, 138);
            this.panelMain.TabIndex = 15;
            this.panelMain.Paint += new System.Windows.Forms.PaintEventHandler(this.panelMain_Paint);
            // 
            // btnRefreshOutlook
            // 
            this.btnRefreshOutlook.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.btnRefreshOutlook.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnRefreshOutlook.FlatAppearance.BorderSize = 0;
            this.btnRefreshOutlook.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRefreshOutlook.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnRefreshOutlook.ForeColor = System.Drawing.Color.White;
            this.btnRefreshOutlook.Location = new System.Drawing.Point(1064, 18);
            this.btnRefreshOutlook.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnRefreshOutlook.Name = "btnRefreshOutlook";
            this.btnRefreshOutlook.Size = new System.Drawing.Size(83, 30);
            this.btnRefreshOutlook.TabIndex = 14;
            this.btnRefreshOutlook.Text = "Yenile";
            this.btnRefreshOutlook.UseVisualStyleBackColor = false;
            this.btnRefreshOutlook.Click += new System.EventHandler(this.btnRefreshOutlook_Click);
            // 
            // mail
            // 
            this.mail.AutoSize = true;
            this.mail.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.mail.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.mail.Location = new System.Drawing.Point(654, 65);
            this.mail.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.mail.Name = "mail";
            this.mail.Size = new System.Drawing.Size(0, 23);
            this.mail.TabIndex = 13;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI Semibold", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label1.Location = new System.Drawing.Point(688, 21);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(161, 23);
            this.label1.TabIndex = 10;
            this.label1.Text = "Kullanılan Hesaplar:";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(1254, 639);
            this.Controls.Add(this.btnOnizleme);
            this.Controls.Add(this.panelMain);
            this.Controls.Add(this.panelTop);
            this.Controls.Add(this.lblHeader);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnSendMails);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MinimumSize = new System.Drawing.Size(799, 499);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Doğum Günü Hatırlatıcı";
            this.Load += new System.EventHandler(this.Form1_Load_1);
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panelTop.ResumeLayout(false);
            this.panelTop.PerformLayout();
            this.panelMain.ResumeLayout(false);
            this.panelMain.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblExcel;
        private System.Windows.Forms.TextBox txtExcelPath;
        private System.Windows.Forms.Button btnSelectExcel;
        private System.Windows.Forms.Button btnSelectFolder;
        private System.Windows.Forms.TextBox txtImageFolder;
        private System.Windows.Forms.Label lblImageFolder;
        private System.Windows.Forms.Button btnSendMails;
        private System.Windows.Forms.ListBox lstLog;
        private System.Windows.Forms.Button btnOnizleme;
        private System.Windows.Forms.ListView LstOnizleme;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label lblHeader;
        private System.Windows.Forms.ColumnHeader columnFirma;
        private System.Windows.Forms.ColumnHeader columnMail;
        private System.Windows.Forms.ColumnHeader columnMudur;
        private System.Windows.Forms.ColumnHeader columnResim;
        private System.Windows.Forms.ColumnHeader columnDogumGunu;
        private System.Windows.Forms.ColumnHeader columnGondericiHesap;
        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.Label lblAppTitle;
        private System.Windows.Forms.Panel panelMain;
        private System.Windows.Forms.Label mail;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnRefreshOutlook;
    }
}

