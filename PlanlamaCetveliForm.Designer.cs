namespace Sinav_Planlama_Cetveli
{
    partial class PlanlamaCetveliForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.dosyaSec = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(242, 252);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Başla";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(228, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Dosya yolu: Lütfen dosyayı buraya sürükleyiniz.";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label2.Location = new System.Drawing.Point(11, 32);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(115, 20);
            this.label2.TabIndex = 2;
            this.label2.Text = "Bilgilendirmeler";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 63);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(274, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Planlama cetvelini buraya sürükleyip bırakabilirsiniz. Veya";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 85);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(482, 13);
            this.label4.TabIndex = 4;
            this.label4.Text = "Planlama cetvelini düzenledikten sonra Excel programını kapatmalısınız. Aksi takd" +
    "irde hata verecektir.";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 107);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(523, 13);
            this.label5.TabIndex = 5;
            this.label5.Text = "Program planlama cetvelinde kesişmeleri kontrol etmeyecektir. Bu nedenle kesişme " +
    "olmadığından emin olunuz.";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(12, 129);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(164, 13);
            this.label6.TabIndex = 6;
            this.label6.Text = "Adayın TCKN\'sini değiştirmeyiniz. ";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(12, 151);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(125, 13);
            this.label7.TabIndex = 7;
            this.label7.Text = "Cetvele satır eklemeyiniz.";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(12, 202);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(516, 13);
            this.label8.TabIndex = 8;
            this.label8.Text = "Planlama cetvelinde düzenleme yaptıktan ve cetveli bulunduran Excel dosyasını kay" +
    "dedip kapattıktan sonra ";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(12, 224);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(365, 13);
            this.label9.TabIndex = 9;
            this.label9.Text = "alt tarafta bulunan başlat butonuna tıklayarak düzenlemeleri kaydedebilirsiniz";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(12, 173);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(274, 13);
            this.label10.TabIndex = 10;
            this.label10.Text = "Programa farklı bir sınavın planlama cetvelini eklemeyiniz.";
            // 
            // dosyaSec
            // 
            this.dosyaSec.Location = new System.Drawing.Point(292, 58);
            this.dosyaSec.Name = "dosyaSec";
            this.dosyaSec.Size = new System.Drawing.Size(75, 23);
            this.dosyaSec.TabIndex = 11;
            this.dosyaSec.Text = "Dosya Seç";
            this.dosyaSec.UseVisualStyleBackColor = true;
            this.dosyaSec.Click += new System.EventHandler(this.dosyaSec_Click);
            // 
            // PlanlamaCetveliForm
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(538, 287);
            this.Controls.Add(this.dosyaSec);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Name = "PlanlamaCetveliForm";
            this.Text = "PlanlamaCetveliForm";
            this.Load += new System.EventHandler(this.PlanlamaCetveliForm_Load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.PlanlamaCetveliForm_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.PlanlamaCetveliForm_DragEnter);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button dosyaSec;
    }
}