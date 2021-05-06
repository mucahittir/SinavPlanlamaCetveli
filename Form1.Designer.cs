namespace Sınav_Planlama_Cetveli
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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.ormanUretimIscisi = new System.Windows.Forms.TabPage();
            this.btnOrmanUretimPlanlamaCetvelindenDuzenle = new System.Windows.Forms.Button();
            this.txtOrmanUretimBaslangicSaati = new System.Windows.Forms.TextBox();
            this.lblSinavAralikOrnek = new System.Windows.Forms.Label();
            this.btnOrmanUretimTabloOlustur = new System.Windows.Forms.Button();
            this.lblBaslangicSaati = new System.Windows.Forms.Label();
            this.lblSinavAralik = new System.Windows.Forms.Label();
            this.lblBaslangicSaatiOrnek = new System.Windows.Forms.Label();
            this.txtOrmanUretimSinavAralik = new System.Windows.Forms.TextBox();
            this.ormanYetistirmeVeBakim = new System.Windows.Forms.TabPage();
            this.txtOrmanYetistirmeVeBakimBaslangicSaati = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnOrmanYetistirmeVeBakimTabloOlustur = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtOrmanYetistirmeVeBakimSinavAralik = new System.Windows.Forms.TextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.lblVersion = new System.Windows.Forms.Label();
            this.btnDosyaSec = new System.Windows.Forms.Button();
            this.ormanUretimData = new System.Windows.Forms.DataGridView();
            this.lblBilgilendirme = new System.Windows.Forms.Label();
            this.lblSayac = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtSinavSuresi = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.txtOrmanYetistirmeSinavSuresi = new System.Windows.Forms.TextBox();
            this.tabControl1.SuspendLayout();
            this.ormanUretimIscisi.SuspendLayout();
            this.ormanYetistirmeVeBakim.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ormanUretimData)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.ormanUretimIscisi);
            this.tabControl1.Controls.Add(this.ormanYetistirmeVeBakim);
            this.tabControl1.Location = new System.Drawing.Point(8, 283);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(349, 202);
            this.tabControl1.TabIndex = 0;
            this.tabControl1.Visible = false;
            // 
            // ormanUretimIscisi
            // 
            this.ormanUretimIscisi.BackColor = System.Drawing.SystemColors.Control;
            this.ormanUretimIscisi.Controls.Add(this.label6);
            this.ormanUretimIscisi.Controls.Add(this.label7);
            this.ormanUretimIscisi.Controls.Add(this.txtSinavSuresi);
            this.ormanUretimIscisi.Controls.Add(this.btnOrmanUretimPlanlamaCetvelindenDuzenle);
            this.ormanUretimIscisi.Controls.Add(this.txtOrmanUretimBaslangicSaati);
            this.ormanUretimIscisi.Controls.Add(this.lblSinavAralikOrnek);
            this.ormanUretimIscisi.Controls.Add(this.btnOrmanUretimTabloOlustur);
            this.ormanUretimIscisi.Controls.Add(this.lblBaslangicSaati);
            this.ormanUretimIscisi.Controls.Add(this.lblSinavAralik);
            this.ormanUretimIscisi.Controls.Add(this.lblBaslangicSaatiOrnek);
            this.ormanUretimIscisi.Controls.Add(this.txtOrmanUretimSinavAralik);
            this.ormanUretimIscisi.Location = new System.Drawing.Point(4, 22);
            this.ormanUretimIscisi.Name = "ormanUretimIscisi";
            this.ormanUretimIscisi.Padding = new System.Windows.Forms.Padding(3);
            this.ormanUretimIscisi.Size = new System.Drawing.Size(341, 176);
            this.ormanUretimIscisi.TabIndex = 0;
            this.ormanUretimIscisi.Text = "Orman Üretim";
            // 
            // btnOrmanUretimPlanlamaCetvelindenDuzenle
            // 
            this.btnOrmanUretimPlanlamaCetvelindenDuzenle.Location = new System.Drawing.Point(49, 147);
            this.btnOrmanUretimPlanlamaCetvelindenDuzenle.Name = "btnOrmanUretimPlanlamaCetvelindenDuzenle";
            this.btnOrmanUretimPlanlamaCetvelindenDuzenle.Size = new System.Drawing.Size(221, 23);
            this.btnOrmanUretimPlanlamaCetvelindenDuzenle.TabIndex = 44;
            this.btnOrmanUretimPlanlamaCetvelindenDuzenle.Text = "Planlama Cetvelinden Saatleri Sisteme Gir";
            this.btnOrmanUretimPlanlamaCetvelindenDuzenle.UseVisualStyleBackColor = true;
            this.btnOrmanUretimPlanlamaCetvelindenDuzenle.Click += new System.EventHandler(this.btnOrmanUretimPlanlamaCetvelindenDuzenle_Click);
            // 
            // txtOrmanUretimBaslangicSaati
            // 
            this.txtOrmanUretimBaslangicSaati.Location = new System.Drawing.Point(147, 24);
            this.txtOrmanUretimBaslangicSaati.Name = "txtOrmanUretimBaslangicSaati";
            this.txtOrmanUretimBaslangicSaati.Size = new System.Drawing.Size(78, 20);
            this.txtOrmanUretimBaslangicSaati.TabIndex = 36;
            this.txtOrmanUretimBaslangicSaati.Text = "08:30";
            this.txtOrmanUretimBaslangicSaati.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // lblSinavAralikOrnek
            // 
            this.lblSinavAralikOrnek.AutoSize = true;
            this.lblSinavAralikOrnek.Location = new System.Drawing.Point(229, 57);
            this.lblSinavAralikOrnek.Name = "lblSinavAralikOrnek";
            this.lblSinavAralikOrnek.Size = new System.Drawing.Size(57, 13);
            this.lblSinavAralikOrnek.TabIndex = 43;
            this.lblSinavAralikOrnek.Text = "Örn: 02:00";
            // 
            // btnOrmanUretimTabloOlustur
            // 
            this.btnOrmanUretimTabloOlustur.Location = new System.Drawing.Point(49, 118);
            this.btnOrmanUretimTabloOlustur.Name = "btnOrmanUretimTabloOlustur";
            this.btnOrmanUretimTabloOlustur.Size = new System.Drawing.Size(221, 23);
            this.btnOrmanUretimTabloOlustur.TabIndex = 35;
            this.btnOrmanUretimTabloOlustur.Text = "Orman Üretim Sınavı Cetveli Oluştur";
            this.btnOrmanUretimTabloOlustur.UseVisualStyleBackColor = true;
            this.btnOrmanUretimTabloOlustur.Click += new System.EventHandler(this.btnOrmanUretimTabloOlustur_Click);
            // 
            // lblBaslangicSaati
            // 
            this.lblBaslangicSaati.AutoSize = true;
            this.lblBaslangicSaati.Location = new System.Drawing.Point(36, 27);
            this.lblBaslangicSaati.Name = "lblBaslangicSaati";
            this.lblBaslangicSaati.Size = new System.Drawing.Size(80, 13);
            this.lblBaslangicSaati.TabIndex = 37;
            this.lblBaslangicSaati.Text = "Başlangıç Saati";
            // 
            // lblSinavAralik
            // 
            this.lblSinavAralik.AutoSize = true;
            this.lblSinavAralik.Location = new System.Drawing.Point(4, 57);
            this.lblSinavAralik.Name = "lblSinavAralik";
            this.lblSinavAralik.Size = new System.Drawing.Size(137, 13);
            this.lblSinavAralik.TabIndex = 42;
            this.lblSinavAralik.Text = "Adayın Sınavları Arası Aralık";
            // 
            // lblBaslangicSaatiOrnek
            // 
            this.lblBaslangicSaatiOrnek.AutoSize = true;
            this.lblBaslangicSaatiOrnek.Location = new System.Drawing.Point(229, 27);
            this.lblBaslangicSaatiOrnek.Name = "lblBaslangicSaatiOrnek";
            this.lblBaslangicSaatiOrnek.Size = new System.Drawing.Size(57, 13);
            this.lblBaslangicSaatiOrnek.TabIndex = 38;
            this.lblBaslangicSaatiOrnek.Text = "Örn: 08:30";
            // 
            // txtOrmanUretimSinavAralik
            // 
            this.txtOrmanUretimSinavAralik.Location = new System.Drawing.Point(147, 54);
            this.txtOrmanUretimSinavAralik.Name = "txtOrmanUretimSinavAralik";
            this.txtOrmanUretimSinavAralik.Size = new System.Drawing.Size(78, 20);
            this.txtOrmanUretimSinavAralik.TabIndex = 41;
            this.txtOrmanUretimSinavAralik.Text = "02:00";
            this.txtOrmanUretimSinavAralik.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // ormanYetistirmeVeBakim
            // 
            this.ormanYetistirmeVeBakim.BackColor = System.Drawing.SystemColors.Control;
            this.ormanYetistirmeVeBakim.Controls.Add(this.label8);
            this.ormanYetistirmeVeBakim.Controls.Add(this.label9);
            this.ormanYetistirmeVeBakim.Controls.Add(this.txtOrmanYetistirmeSinavSuresi);
            this.ormanYetistirmeVeBakim.Controls.Add(this.button1);
            this.ormanYetistirmeVeBakim.Controls.Add(this.txtOrmanYetistirmeVeBakimBaslangicSaati);
            this.ormanYetistirmeVeBakim.Controls.Add(this.label2);
            this.ormanYetistirmeVeBakim.Controls.Add(this.btnOrmanYetistirmeVeBakimTabloOlustur);
            this.ormanYetistirmeVeBakim.Controls.Add(this.label3);
            this.ormanYetistirmeVeBakim.Controls.Add(this.label4);
            this.ormanYetistirmeVeBakim.Controls.Add(this.label5);
            this.ormanYetistirmeVeBakim.Controls.Add(this.txtOrmanYetistirmeVeBakimSinavAralik);
            this.ormanYetistirmeVeBakim.Location = new System.Drawing.Point(4, 22);
            this.ormanYetistirmeVeBakim.Name = "ormanYetistirmeVeBakim";
            this.ormanYetistirmeVeBakim.Padding = new System.Windows.Forms.Padding(3);
            this.ormanYetistirmeVeBakim.Size = new System.Drawing.Size(341, 176);
            this.ormanYetistirmeVeBakim.TabIndex = 1;
            this.ormanYetistirmeVeBakim.Text = "Orman Yetiştirme ve Bakım";
            // 
            // txtOrmanYetistirmeVeBakimBaslangicSaati
            // 
            this.txtOrmanYetistirmeVeBakimBaslangicSaati.Location = new System.Drawing.Point(147, 24);
            this.txtOrmanYetistirmeVeBakimBaslangicSaati.Name = "txtOrmanYetistirmeVeBakimBaslangicSaati";
            this.txtOrmanYetistirmeVeBakimBaslangicSaati.Size = new System.Drawing.Size(78, 20);
            this.txtOrmanYetistirmeVeBakimBaslangicSaati.TabIndex = 45;
            this.txtOrmanYetistirmeVeBakimBaslangicSaati.Text = "08:30";
            this.txtOrmanYetistirmeVeBakimBaslangicSaati.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(229, 53);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(57, 13);
            this.label2.TabIndex = 50;
            this.label2.Text = "Örn: 02:00";
            // 
            // btnOrmanYetistirmeVeBakimTabloOlustur
            // 
            this.btnOrmanYetistirmeVeBakimTabloOlustur.Location = new System.Drawing.Point(49, 114);
            this.btnOrmanYetistirmeVeBakimTabloOlustur.Name = "btnOrmanYetistirmeVeBakimTabloOlustur";
            this.btnOrmanYetistirmeVeBakimTabloOlustur.Size = new System.Drawing.Size(215, 23);
            this.btnOrmanYetistirmeVeBakimTabloOlustur.TabIndex = 44;
            this.btnOrmanYetistirmeVeBakimTabloOlustur.Text = "Orman Yetiştirme Sınavı Cetveli Oluştur";
            this.btnOrmanYetistirmeVeBakimTabloOlustur.UseVisualStyleBackColor = true;
            this.btnOrmanYetistirmeVeBakimTabloOlustur.Click += new System.EventHandler(this.btnOrmanYetistirmeVeBakimTabloOlustur_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(36, 27);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(80, 13);
            this.label3.TabIndex = 46;
            this.label3.Text = "Başlangıç Saati";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(4, 53);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(137, 13);
            this.label4.TabIndex = 49;
            this.label4.Text = "Adayın Sınavları Arası Aralık";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(229, 27);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(57, 13);
            this.label5.TabIndex = 47;
            this.label5.Text = "Örn: 08:30";
            // 
            // txtOrmanYetistirmeVeBakimSinavAralik
            // 
            this.txtOrmanYetistirmeVeBakimSinavAralik.Location = new System.Drawing.Point(147, 50);
            this.txtOrmanYetistirmeVeBakimSinavAralik.Name = "txtOrmanYetistirmeVeBakimSinavAralik";
            this.txtOrmanYetistirmeVeBakimSinavAralik.Size = new System.Drawing.Size(78, 20);
            this.txtOrmanYetistirmeVeBakimSinavAralik.TabIndex = 48;
            this.txtOrmanYetistirmeVeBakimSinavAralik.Text = "02:00";
            this.txtOrmanYetistirmeVeBakimSinavAralik.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 504);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(341, 22);
            this.progressBar1.TabIndex = 46;
            this.progressBar1.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(58, 212);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(234, 13);
            this.label1.TabIndex = 45;
            this.label1.Text = "(Aday bildirim dosyanızı sürükleyip bırakabilirsiniz)";
            // 
            // webBrowser1
            // 
            this.webBrowser1.AllowWebBrowserDrop = false;
            this.webBrowser1.Location = new System.Drawing.Point(8, 12);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(341, 183);
            this.webBrowser1.TabIndex = 44;
            // 
            // lblVersion
            // 
            this.lblVersion.AutoSize = true;
            this.lblVersion.Location = new System.Drawing.Point(312, 529);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(41, 13);
            this.lblVersion.TabIndex = 40;
            this.lblVersion.Text = "asdasd";
            // 
            // btnDosyaSec
            // 
            this.btnDosyaSec.Location = new System.Drawing.Point(130, 245);
            this.btnDosyaSec.Name = "btnDosyaSec";
            this.btnDosyaSec.Size = new System.Drawing.Size(103, 23);
            this.btnDosyaSec.TabIndex = 34;
            this.btnDosyaSec.Text = "Aday Bildirim";
            this.btnDosyaSec.UseVisualStyleBackColor = true;
            this.btnDosyaSec.Click += new System.EventHandler(this.btnDosyaSec_Click);
            // 
            // ormanUretimData
            // 
            this.ormanUretimData.AllowDrop = true;
            this.ormanUretimData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ormanUretimData.Location = new System.Drawing.Point(8, 12);
            this.ormanUretimData.Name = "ormanUretimData";
            this.ormanUretimData.Size = new System.Drawing.Size(341, 183);
            this.ormanUretimData.TabIndex = 33;
            this.ormanUretimData.DragDrop += new System.Windows.Forms.DragEventHandler(this.dataGridView1_DragDrop);
            this.ormanUretimData.DragEnter += new System.Windows.Forms.DragEventHandler(this.dataGridView1_DragEnter);
            // 
            // lblBilgilendirme
            // 
            this.lblBilgilendirme.AutoSize = true;
            this.lblBilgilendirme.Location = new System.Drawing.Point(40, 488);
            this.lblBilgilendirme.Name = "lblBilgilendirme";
            this.lblBilgilendirme.Size = new System.Drawing.Size(242, 13);
            this.lblBilgilendirme.TabIndex = 47;
            this.lblBilgilendirme.Text = "Kullanıcının Sayemyis\'e Giriş Yapması Bekleniyor...";
            this.lblBilgilendirme.Visible = false;
            // 
            // lblSayac
            // 
            this.lblSayac.AutoSize = true;
            this.lblSayac.Location = new System.Drawing.Point(282, 488);
            this.lblSayac.Name = "lblSayac";
            this.lblSayac.Size = new System.Drawing.Size(12, 13);
            this.lblSayac.TabIndex = 48;
            this.lblSayac.Text = "x";
            this.lblSayac.Visible = false;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(49, 143);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(215, 23);
            this.button1.TabIndex = 51;
            this.button1.Text = "Planlama Cetvelinden Saatleri Sisteme Gir";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(229, 84);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(42, 13);
            this.label6.TabIndex = 47;
            this.label6.Text = "Örn: 30";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(19, 84);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(109, 13);
            this.label7.TabIndex = 46;
            this.label7.Text = "Sınav Süresi (Dakika)";
            // 
            // txtSinavSuresi
            // 
            this.txtSinavSuresi.Location = new System.Drawing.Point(147, 81);
            this.txtSinavSuresi.Name = "txtSinavSuresi";
            this.txtSinavSuresi.Size = new System.Drawing.Size(78, 20);
            this.txtSinavSuresi.TabIndex = 45;
            this.txtSinavSuresi.Text = "30";
            this.txtSinavSuresi.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(229, 79);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(42, 13);
            this.label8.TabIndex = 54;
            this.label8.Text = "Örn: 30";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(28, 79);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(109, 13);
            this.label9.TabIndex = 53;
            this.label9.Text = "Sınav Süresi (Dakika)";
            // 
            // txtOrmanYetistirmeSinavSuresi
            // 
            this.txtOrmanYetistirmeSinavSuresi.Location = new System.Drawing.Point(147, 76);
            this.txtOrmanYetistirmeSinavSuresi.Name = "txtOrmanYetistirmeSinavSuresi";
            this.txtOrmanYetistirmeSinavSuresi.Size = new System.Drawing.Size(78, 20);
            this.txtOrmanYetistirmeSinavSuresi.TabIndex = 52;
            this.txtOrmanYetistirmeSinavSuresi.Text = "30";
            this.txtOrmanYetistirmeSinavSuresi.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // Form1
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(357, 551);
            this.Controls.Add(this.lblSayac);
            this.Controls.Add(this.lblBilgilendirme);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.webBrowser1);
            this.Controls.Add(this.lblVersion);
            this.Controls.Add(this.ormanUretimData);
            this.Controls.Add(this.btnDosyaSec);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Sınav Planlama Cetveli";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.dataGridView1_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.dataGridView1_DragEnter);
            this.tabControl1.ResumeLayout(false);
            this.ormanUretimIscisi.ResumeLayout(false);
            this.ormanUretimIscisi.PerformLayout();
            this.ormanYetistirmeVeBakim.ResumeLayout(false);
            this.ormanYetistirmeVeBakim.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ormanUretimData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage ormanUretimIscisi;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.Label lblSinavAralikOrnek;
        private System.Windows.Forms.Label lblSinavAralik;
        private System.Windows.Forms.TextBox txtOrmanUretimSinavAralik;
        private System.Windows.Forms.Label lblVersion;
        private System.Windows.Forms.Label lblBaslangicSaatiOrnek;
        private System.Windows.Forms.Label lblBaslangicSaati;
        private System.Windows.Forms.TextBox txtOrmanUretimBaslangicSaati;
        private System.Windows.Forms.Button btnOrmanUretimTabloOlustur;
        private System.Windows.Forms.Button btnDosyaSec;
        private System.Windows.Forms.DataGridView ormanUretimData;
        private System.Windows.Forms.TabPage ormanYetistirmeVeBakim;
        private System.Windows.Forms.TextBox txtOrmanYetistirmeVeBakimBaslangicSaati;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnOrmanYetistirmeVeBakimTabloOlustur;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtOrmanYetistirmeVeBakimSinavAralik;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblBilgilendirme;
        private System.Windows.Forms.Label lblSayac;
        private System.Windows.Forms.Button btnOrmanUretimPlanlamaCetvelindenDuzenle;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtSinavSuresi;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtOrmanYetistirmeSinavSuresi;
    }
}

