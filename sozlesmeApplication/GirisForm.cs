using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace sozlesmeApplication
{
    public partial class GirisForm : Form
    {
        public GirisForm()
        {
            InitializeComponent();
        }

        public static string kullanici;

        private void girisYapButton_Click(object sender, EventArgs e)
        {
            if (kullaniciSecimComboBox.Text != "")
            {
                Form1 form1 = new Form1();
                this.Hide();
                form1.Show();

                kullanici = kullaniciSecimComboBox.Text;
            }
            else
            {
                MessageBox.Show("Lütfen kullanıcı seçiniz.");
            }
        }

        private void GirisForm_Load(object sender, EventArgs e)
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            filePath += @"\Update For Sozlesme.exe";

            try
            {
                Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Güncelleme kontrol edilirken hata oluştu: " + ex.Message);
            }
        }
    }
}
