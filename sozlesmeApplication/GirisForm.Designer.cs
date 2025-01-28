namespace sozlesmeApplication
{
    partial class GirisForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.kullaniciSecimComboBox = new System.Windows.Forms.ComboBox();
            this.girisYapButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(81, 57);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 13);
            this.label1.TabIndex = 19;
            this.label1.Text = "Kullanıcı seçiniz:";
            // 
            // kullaniciSecimComboBox
            // 
            this.kullaniciSecimComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.kullaniciSecimComboBox.FormattingEnabled = true;
            this.kullaniciSecimComboBox.Items.AddRange(new object[] {
            "Furkan Hoşgör",
            "Hasan Uysal",
            "Hayati Uçan",
            "Taner Varsak",
            "Semih Aygün",
            "Samet Adalı"});
            this.kullaniciSecimComboBox.Location = new System.Drawing.Point(81, 75);
            this.kullaniciSecimComboBox.Margin = new System.Windows.Forms.Padding(2);
            this.kullaniciSecimComboBox.Name = "kullaniciSecimComboBox";
            this.kullaniciSecimComboBox.Size = new System.Drawing.Size(92, 21);
            this.kullaniciSecimComboBox.TabIndex = 18;
            // 
            // girisYapButton
            // 
            this.girisYapButton.Location = new System.Drawing.Point(98, 128);
            this.girisYapButton.Margin = new System.Windows.Forms.Padding(2);
            this.girisYapButton.Name = "girisYapButton";
            this.girisYapButton.Size = new System.Drawing.Size(56, 45);
            this.girisYapButton.TabIndex = 17;
            this.girisYapButton.Text = "Giriş Yap";
            this.girisYapButton.UseVisualStyleBackColor = true;
            this.girisYapButton.Click += new System.EventHandler(this.girisYapButton_Click);
            // 
            // GirisForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(251, 221);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.kullaniciSecimComboBox);
            this.Controls.Add(this.girisYapButton);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "GirisForm";
            this.Text = "Giriş Paneli";
            this.Load += new System.EventHandler(this.GirisForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox kullaniciSecimComboBox;
        private System.Windows.Forms.Button girisYapButton;
    }
}