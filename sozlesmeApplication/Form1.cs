using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using static System.Net.Mime.MediaTypeNames;
using Application = Microsoft.Office.Interop.Word.Application;
using Rectangle = System.Drawing.Rectangle;

namespace sozlesmeApplication
{
    public partial class Form1 : Form
    {
        private readonly SozlesmeRepository _sozlesmeRepository;
        string connectionString = $"Data Source={DatabaseConnectionHelper.GetServerName()};Initial Catalog=sozlesme;Integrated Security=True;Connect Timeout=30;Encrypt=True;TrustServerCertificate=True";

        List<Label> lblList = new List<Label>();
        List<TextBox> txtList = new List<TextBox>();
        List<TextBox> txtTasinmazList = new List<TextBox>();
        //List<(int, Control)> values = new List<(int, Control)>();

        private List<List<TextBox>> textBoxList;

        List<DateTimePicker> dtpList = new List<DateTimePicker>();

        Microsoft.Office.Interop.Word.Application wordApp;
        //Document wordDoc;

        List<Control> pesinComponentList = new List<Control>();
        List<Control> takasComponentList = new List<Control>();
        List<Control> taksitComponentList = new List<Control>();
        List<Control> taksitTutariComponentList = new List<Control>();

        public Form1()
        {
            InitializeComponent();

            DatabaseConnectionHelper.SetupDatabase();
            DatabaseHelper databaseHelper = new DatabaseHelper(connectionString);
            _sozlesmeRepository = new SozlesmeRepository(databaseHelper);

            foreach (var item in this.Controls)
            {
                if (item is DateTimePicker dateTimePicker)
                {
                    // Format ve CustomFormat ayarlarını yapıyoruz
                    dateTimePicker.Format = DateTimePickerFormat.Short;
                    //dateTimePicker.CustomFormat = "dd/MM/yyyy"; // Tarih formatı
                }
            }

            pesinComponentList.Add(pesinOdemeTutariTextBox);
            pesinComponentList.Add(pesinOdemeTarihiDateTimePicker);

            takasComponentList.Add(takasMalCinsTextBox);
            takasComponentList.Add(takasMalOzellikTextBox);
            takasComponentList.Add(takasMalTutarTextBox);
            takasComponentList.Add(takasMalTeslimTarihiDateTimePicker);

            taksitComponentList.Add(taksitSayisiComboBox);
            taksitComponentList.Add(taksitBaslangicTarihiDateTimePicker);
            taksitComponentList.Add(taksitTutariTextBox);
        }

        Dictionary<string, (string TextBoxValue, DateTime DateTimePickerValue)> data = new Dictionary<string, (string, DateTime)>();

        // Birler basamağı için kelimeler
        static string[] units = { "", "bir", "iki", "üç", "dört", "beş", "altı", "yedi", "sekiz", "dokuz" };
        // Onlar basamağı için kelimeler
        static string[] tens = { "", "on", "yirmi", "otuz", "kırk", "elli", "altmış", "yetmiş", "seksen", "doksan" };
        // Büyük sayılar için ekler
        static string[] scales = { "", "bin", "milyon", "milyar", "trilyon" };

        public static string NumberToCurrencyText(decimal number)
        {
            if (number == 0)
                return "sıfırtürklirası";

            // Sayının tam kısmını ve kuruş kısmını ayırıyoruz
            long lira = (long)Math.Floor(number);
            int kurus = (int)((number - lira) * 100);

            string result = "";

            bool z = false;
            // Tam kısmı yazıya çeviriyoruz
            if (kurus > 0)
            {
                z = true;
            }

            if (lira > 0)
            {
                if (z)
                {
                    result += NumberToText(lira) + "türklirası";
                }
                else
                {
                    result += NumberToText(lira) + "türklirası)";
                }
            }

            // Kuruş kısmı varsa onu da yazıya çeviriyoruz
            if (kurus > 0)
            {
                result += NumberToText(kurus) + "kuruş)";
            }

            return result;
        }

        public static string NumberToText(long number)
        {
            if (number == 0)
                return "sıfır";

            string numberStr = number.ToString();
            // Sayının uzunluğunu 3'ün katı olacak şekilde başına sıfır ekleyerek düzenliyoruz
            int numLength = numberStr.Length;
            int mod = numLength % 3;
            if (mod != 0)
                numberStr = numberStr.PadLeft(numLength + (3 - mod), '0');

            int totalGroups = numberStr.Length / 3;
            string result = "";
            for (int i = 0; i < totalGroups; i++)
            {
                int groupValue = int.Parse(numberStr.Substring(i * 3, 3));
                int scaleIndex = totalGroups - i - 1;
                if (groupValue != 0)
                {
                    string groupText = ConvertGroup(groupValue);
                    if (scaleIndex == 1 && groupValue == 1)
                    {
                        // Özel durum: "birbin" yerine sadece "bin" kullanılır
                        result += scales[scaleIndex];
                    }
                    else
                    {
                        result += groupText + scales[scaleIndex];
                    }
                }
            }
            return result;
        }

        static string ConvertGroup(int number)
        {
            string result = "";

            int hundreds = number / 100;
            int remainder = number % 100;

            // Yüzler basamağı
            if (hundreds != 0)
            {
                if (hundreds == 1)
                {
                    result += "yüz";
                }
                else
                {
                    result += units[hundreds] + "yüz";
                }
            }

            // Onlar basamağı
            int tensPlace = remainder / 10;
            int unitsPlace = remainder % 10;

            if (tensPlace != 0)
            {
                result += tens[tensPlace];
            }

            // Birler basamağı
            if (unitsPlace != 0)
            {
                result += units[unitsPlace];
            }

            return result;
        }

        object missing = System.Reflection.Missing.Value;
        //Microsoft.Office.Interop.Word.Application app;
        private void button1_Click(object sender, EventArgs e)
        {
            klasorOlustur();
            belgeUret();
        }


        // Sözleşme Ekle
        public void SozlesmeEkle(Sozlesme sozlesme)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = @"INSERT INTO sozlesme (sozlesmeyiOlusturanKullanici, saticiUnvan, aliciTc, aliciVN, tasinmazBedeli, pesinOdemeTutari, pesinOdemeTarihi,
                              takasMalinCinsi, takasMalinOzellikleri, takasMalinTutari, takasMalinTeslimTarihi,
                              taksitSayisi, taksitTutari, taksitBaslangicTarihi, taksitBitisTarihi, sozlesmeTarihi)
                             VALUES ( @sozlesmeyiOlusturanKullanici, @saticiUnvan, @aliciTc, @aliciVN, @tasinmazBedeli, @pesinOdemeTutari, @pesinOdemeTarihi,
                              @takasMalinCinsi, @takasMalinOzellikleri, @takasMalinTutari, @takasMalinTeslimTarihi,
                              @taksitSayisi, @taksitTutari, @taksitBaslangicTarihi, @taksitBitisTarihi, @sozlesmeTarihi)";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@sozlesmeyiOlusturanKullanici", sozlesme.SozlesmeyiOlusturanKullanici);
                    cmd.Parameters.AddWithValue("@saticiUnvan", sozlesme.SaticiUnvan);
                    cmd.Parameters.AddWithValue("@sozlesmeTarihi", sozlesme.SozlesmeTarihi);
                    // Nullable alanları kontrol ederek ekleme
                    cmd.Parameters.AddWithValue("@aliciTc", (object?)sozlesme.AliciTc ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@aliciVN", (object?)sozlesme.AliciVN ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@tasinmazBedeli", (object?)sozlesme.TasinmazBedeli ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@pesinOdemeTutari", (object?)sozlesme.PesinOdemeTutari ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@pesinOdemeTarihi", (object?)sozlesme.PesinOdemeTarihi ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@takasMalinCinsi", (object?)sozlesme.TakasMalinCinsi ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@takasMalinOzellikleri", (object?)sozlesme.TakasMalinOzellikleri ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@takasMalinTutari", (object?)sozlesme.TakasMalinTutari ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@takasMalinTeslimTarihi", (object?)sozlesme.TakasMalinTeslimTarihi ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@taksitSayisi", (object?)sozlesme.TaksitSayisi ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@taksitTutari", (object?)sozlesme.TaksitTutari ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@taksitBaslangicTarihi", (object?)sozlesme.TaksitBaslangicTarihi ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@taksitBitisTarihi", (object?)sozlesme.TaksitBitisTarihi ?? DBNull.Value);

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // Bireysel Alıcı Ekle
        public void BireyselAliciEkle(AliciBireysel alici)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = "INSERT INTO alici_bireysel (aliciTc, adSoyad, aliciAdres, aliciTelefon) " +
                               "VALUES (@aliciTc, @adSoyad, @aliciAdres, @aliciTelefon)";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@aliciTc", alici.AliciTc);
                    cmd.Parameters.AddWithValue("@adSoyad", alici.AdSoyad);
                    cmd.Parameters.AddWithValue("@aliciAdres", alici.AliciAdres);
                    cmd.Parameters.AddWithValue("@aliciTelefon", alici.AliciTelefon);

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
        }



        // Şirket Alıcı Ekle
        public void AliciSirketEkle(AliciSirket sirket)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = "INSERT INTO alici_sirket (vergiNumarasi, vergiDairesi, sirketAdres, sirketTelefon, sirketUnvan) " +
                               "VALUES (@vergiNumarasi, @vergiDairesi, @sirketAdres, @sirketTelefon, @sirketUnvan)";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@vergiNumarasi", sirket.VergiNumarasi);
                    cmd.Parameters.AddWithValue("@vergiDairesi", sirket.VergiDairesi);
                    cmd.Parameters.AddWithValue("@sirketAdres", sirket.SirketAdres);
                    cmd.Parameters.AddWithValue("@sirketTelefon", sirket.SirketTelefon);
                    cmd.Parameters.AddWithValue("@sirketUnvan", sirket.SirketUnvan);

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public async void yeniSozlesme()
        {
            if (saticiComboBox.Text == "")
            {
                MessageBox.Show("Satıcı seçilmedi!");
                return;
            }
            if (telTextBox.Text.Length != 10)
            {
                MessageBox.Show("Telefon numarası yanlış veya eksik girildi.");
                return;
            }
            if (adresTextBox.Text == "")
            {
                MessageBox.Show("Adres kısmı boş geçilemez!");
                return;
            }

            int aliciBireyselCount = 0, aliciSirketCount = 0;
            foreach (var item in aliciBireyselComponentList)
            {
                if (item.Text == "")
                {
                    aliciBireyselCount++;
                }
            }

            foreach (var item in aliciSirketComponentList)
            {
                if (item.Text == "")
                {
                    aliciSirketCount++;
                }
            }
            if (aliciBireyselCount > 0 && aliciSirketCount > 0)
            {
                MessageBox.Show("Taşınmaz sahibi bilgileri eksik veya yanlış girildi. Lütfen kontrol ediniz.");
                return;
            }
            string aliciAdSoyadUnvan = "";
            if (unvanTextBox.Text == "")
            {
                aliciAdSoyadUnvan = adSoyadTextBox.Text;
            }
            else
            {
                aliciAdSoyadUnvan = unvanTextBox.Text;
            }
            string aliciTcnVn = "";

            if (vnTextBox.Text == "")
            {
                try
                {
                    aliciTcnVn = tcknTextBox.Text.Substring(0, 11);
                }
                catch (ArgumentOutOfRangeException)
                {
                    MessageBox.Show("Lütfen TC kimlik numarasının doğru girildiğinden emin olunuz.");
                }
            }
            else if (tcknTextBox.Text == "")
            {
                try
                {
                    aliciTcnVn = vergiDairesiTextBox.Text + "/" + vnTextBox.Text.Substring(0, 10);
                }
                catch (ArgumentOutOfRangeException)
                {
                    MessageBox.Show("Lütfen vergi numarasının doğru girildiğinden emin olunuz.");
                }
            }

            string aliciAdres = adresTextBox.Text;
            string aliciTel = telTextBox.Text;

            string sUnvan = "", sVDVN = "", sMersis = "", sAdres = "";
            if (saticiComboBox.Text == "Bereketli Topraklar")
            {
                sUnvan = "BEREKETLİ EMLAK SANAYİ VE TİCARET ANONİM ŞİRKETİ";
                sVDVN = "OSMANGAZİ/ 1650764758";
                sMersis = "0165076475800001";
                sAdres = "NİLÜFERKÖY MAH. NİLÜFER DEĞİRMEN SK. A BLOK NO: 12/7                                               Osmangazi/BURSA";
            }
            else if (saticiComboBox.Text == "BT Meta İnşaat")
            {
                sUnvan = "BT META İNŞAAT SANAYİ VE TİCARET ANONİM ŞİRKETİ";
                sVDVN = "KURTDERELİ/ 1871662733";
                sMersis = "0187166273300001";
                sAdres = "ÇAYIRHİSAR MAH. 6000 SK. NO: 6 İÇ KAPI NO: 3 ALTIEYLÜL/BALIKESİR";
            }
            else if (saticiComboBox.Text == "Bilal Aktaş")
            {
                sUnvan = "BİLAL AKTAŞ";
                sVDVN = "Nilüfer/ 2322798326";
                sMersis = "";
                sAdres = "İhsaniye mah. Leylak sk. No: 3B Daire: 13 Nilüfer/Bursa";
            }

            if (tcknTextBox.Text != "" && vnTextBox.Text != "")
            {
                MessageBox.Show("Aynı anda hem şirket hem de bireysel alıcı bilgisi girilemez!");
                return;
            }

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = null;

            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Word", "ihtimaller");

            filePath = filePath + @"\yetkilendirmeSatis.docx";

            if (IsFileLocked(filePath))
            {
                MessageBox.Show("Dosya açık veya kilitli. Lütfen Kill All işlevini kullandıktan sonra yeniden deneyiniz.");
                return;
            }

            doc = wordApp.Documents.Open(filePath);

            int tasinmazSayisi;
            if (int.TryParse(tasinmazSayisiComboBox.Text, out tasinmazSayisi))
            {

            }
            else
            {
                MessageBox.Show("Taşınmaz sayısı doğru biçimde girilmedi.");
                return;
            }

            if (tasinmazSayisi % 2 == 0)
            {
                AddTablesBelowSentence(doc, "MADDE 2: SÖZLEŞMEYE KONU TAŞINMAZ", tasinmazSayisi / 2, 9, 4);
            }
            else if (tasinmazSayisi % 2 == 1)
            {
                AddTablesBelowSentence(doc, "MADDE 2: SÖZLEŞMEYE KONU TAŞINMAZ", (tasinmazSayisi + 1) / 2, 9, 4);
            }
            klasorOlustur();
            // alıcı ve satıcı bilgilerini giren döngü
            foreach (ContentControl control in doc.ContentControls)
            {
                if (control.Tag == "sozlesmeTarihiText")
                {
                    control.Range.Text = sozlesmeTarihiDateTimePicker.Value.ToString("dd'/'MM'/'yyyy"); ;
                }
                if (control.Tag == "sUnvanText")
                {
                    control.Range.Text = sUnvan;
                }
                if (control.Tag == "sVDVNText")
                {
                    control.Range.Text = sVDVN;
                }
                if (control.Tag == "sMersisText")
                {
                    if (sMersis == "")
                    {
                        control.Range.Text = ".";
                        control.Range.Font.Size = 1;
                    }
                    else
                    {
                        control.Range.Text = sMersis;
                        control.Range.Font.Size = 11;
                    }
                }
                if (control.Tag == "sAdresText")
                {
                    control.Range.Text = sAdres;
                }

                if (control.Tag == "aAdSoyadUnvanText")
                {
                    control.Range.Text = aliciAdSoyadUnvan;
                }
                if (control.Tag == "aTCNVNText")
                {
                    control.Range.Text = aliciTcnVn;
                }
                if (control.Tag == "aAdresText")
                {
                    control.MultiLine = true; // İçerik denetimini çok satırlı yap
                    control.Range.Text = aliciAdres;
                }
                if (control.Tag == "aTelefonText")
                {
                    control.Range.Text = aliciTel;
                }
            }

            foreach (Section section in doc.Sections)
            {
                foreach (HeaderFooter footer in section.Footers)
                {
                    foreach (ContentControl footerControl in footer.Range.ContentControls)
                    {
                        if (footerControl.Tag == "sVDVNText")
                        {
                            footerControl.Range.Text = sVDVN;
                        }
                        if (footerControl.Tag == "sUnvanText")
                        {
                            footerControl.Range.Text = sUnvan;
                        }
                        if (footerControl.Tag == "aAdSoyadUnvanText")
                        {
                            footerControl.Range.Text = aliciAdSoyadUnvan;
                        }
                    }
                }
            }

            try
            {
                string tarih = DateTime.Today.ToString("dd-MM-yyyy");
                string fileName = aliciAdSoyadUnvan + " - " + tarih + " - Yetkilendirme Sözleşmesi.docx";

                doc.SaveAs2(Path.Combine(kullanilanSozlesmelerYolu, fileName));
                doc.SaveAs2(Path.Combine(sozlesmeArsiviYolu, fileName));
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("Dosya kaydedilirken hata: Lütfen 'Kill All' ve 'Temizle' fonksiyonlarını kullandıktan sonra yeniden deneyiniz.");
                return;
            }

            try
            {
                doc?.Close(ref missing, ref missing, ref missing);
                wordApp.Application.Quit(ref missing, ref missing, ref missing);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);

                await System.Threading.Tasks.Task.Delay(1000);
                notifyIcon1.Icon = SystemIcons.Information; // Veya kendi .ico dosyanız
                notifyIcon1.Visible = true;
                notifyIcon1.Text = "Sözleşme Üretimi";
                notifyIcon1.BalloonTipTitle = "Belge Üretimi";
                notifyIcon1.BalloonTipText = "Belge üretimi başarıyla tamamlandı.";
                notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
                notifyIcon1.ShowBalloonTip(5000);
                await System.Threading.Tasks.Task.Delay(5500);
                notifyIcon1.Visible = false;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("İşlem başarısız, lütfen 'Temizle' ve 'Kill All' butonlarını kullandıktan sonra tekrar deneyiniz.");
                throw;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }



        public async void belgeUret()
        {
            decimal pesinOdemeTutari;
            decimal takasMalTutari;
            int count = 0;

            bool pesin = false, takas = false, taksit = false;

            if (saticiComboBox.Text == "")
            {
                MessageBox.Show("Satıcı seçilmedi!");
                return;
            }

            if (telTextBox.Text.Length != 10)
            {
                MessageBox.Show("Telefon numarası yanlış veya eksik girildi.");
                return;
            }
            if (adresTextBox.Text.Length > 0)
            {

            }
            else
            {
                MessageBox.Show("Adres kısmı boş geçilemez!");
                return;
            }


            int aliciBireyselCount = 0, aliciSirketCount = 0;
            foreach (var item in aliciBireyselComponentList)
            {
                if (item.Text == "")
                {
                    aliciBireyselCount++;
                }
            }

            foreach (var item in aliciSirketComponentList)
            {
                if (item.Text == "")
                {
                    aliciSirketCount++;
                }
            }
            if (aliciBireyselCount > 0 && aliciSirketCount > 0)
            {
                MessageBox.Show("Alıcı bilgileri eksik. Lütfen kontrol ediniz.");
                return;
            }

            if (decimal.TryParse(pesinOdemeTutariTextBox.Text, out pesinOdemeTutari))
            {
                pesin = true;
            }
            else
            {
                if (pesinCheckBox.Checked)
                {
                    MessageBox.Show("Peşin tutarı doğru biçimde girilmedi!");
                    return;
                }
                count++;
            }
            if (decimal.TryParse(takasMalTutarTextBox.Text, out takasMalTutari))
            {
                takas = true;
            }
            else
            {
                if (takasCheckBox.Checked)
                {
                    MessageBox.Show("Takas tutarı doğru biçimde girilmedi!");
                    return;
                }
                count++;
            }

            int taksitSayisi;
            if (int.TryParse(taksitSayisiComboBox.Text, out taksitSayisi))
            {

            }
            else
            {
                if (taksitCheckBox.Checked)
                {
                    MessageBox.Show("Taksit sayısı doğru biçimde girilmedi!");
                    return;
                }
                count++;
            }


            if (count == 3)
            {
                MessageBox.Show("Peşin, takas veya taksit tutarlarından en az biri doğru biçimde girilmedi!");
                return;
            }

            int tasinmazSayisi;
            if (int.TryParse(tasinmazSayisiComboBox.Text, out tasinmazSayisi))
            {

            }
            else
            {
                MessageBox.Show("Taşınmaz sayısı doğru biçimde girilmedi.");
                return;
            }



            bool pesinC = pesinCheckBox.Checked;
            bool takasC = takasCheckBox.Checked;
            bool taksitC = taksitCheckBox.Checked;

            int ihtimal = 0;
            //her bir ihtimale göre farklı bir doküman açılacak
            //pesin, takas ve taksit
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Word", "ihtimaller");
            if (pesinC && takasC && taksitC)
            {
                filePath = filePath + @"\pesinTakasTaksit.docx";
                ihtimal = 1;
            }//pesin ve takas 
            else if (pesinC && takasC && !taksitC)
            {
                filePath = filePath + @"\pesinTakas.docx";
                ihtimal = 2;
            }//pesin ve taksit
            else if (pesinC && taksitC && !takasC)
            {
                filePath = filePath + @"\pesinTaksit.docx";
                ihtimal = 3;
            }//takas ve taksit
            else if (takasC && taksitC && !pesinC)
            {
                filePath = filePath + @"\takasTaksit.docx";
                ihtimal = 4;
            }//sadece pesin
            else if (pesinC && !takasC && !taksitC)
            {
                filePath = filePath + @"\pesin.docx";
                ihtimal = 5;
            }//sadece takas
            else if (takasC && !taksitC && !pesinC)
            {
                filePath = filePath + @"\takas.docx";
                ihtimal = 6;
            }//sadece taksit
            else if (taksitC && !takasC && !pesinC)
            {
                filePath = filePath + @"\taksit.docx";
                ihtimal = 7;
            }

            decimal taksitTutari = 0;
            foreach (var item in taksitTutariComponentList)
            {
                if (item is TextBox txt2)
                {
                    decimal taksitTutariTextBox;
                    if (decimal.TryParse(txt2.Text, out taksitTutariTextBox))
                    {
                        taksit = true;
                        taksitTutari += decimal.Parse(txt2.Text);
                    }
                }
            }


            decimal tasinmazBedeli = 0;

            if (pesin)
            {
                tasinmazBedeli += pesinOdemeTutari;
            }
            if (takas)
            {
                tasinmazBedeli += takasMalTutari;
            }
            if (taksit)
            {
                tasinmazBedeli += taksitTutari;
            }

            if (filePath == "")
            {
                MessageBox.Show("Peşin, taksit veya takas seçeneklerinden en az biri seçilmelidir!");
                return;
            }
            if (IsFileLocked(filePath))
            {
                MessageBox.Show("Dosya açık veya kilitli. Lütfen Kill All işlevini kullandıktan sonra yeniden deneyiniz.");
                return;
            }

            string tempFilePath = Path.Combine(Path.GetDirectoryName(filePath), "~$" + Path.GetFileName(filePath));
            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }

            //string wordPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Word");

            string aliciAdSoyadUnvan = "";
            if (unvanTextBox.Text == "")
            {
                aliciAdSoyadUnvan = adSoyadTextBox.Text;
            }
            else
            {
                aliciAdSoyadUnvan = unvanTextBox.Text;
            }
            string aliciTcnVn = "";

            if (vnTextBox.Text == "")
            {
                try
                {
                    aliciTcnVn = tcknTextBox.Text.Substring(0, 11);
                }
                catch (ArgumentOutOfRangeException)
                {
                    MessageBox.Show("Lütfen TC kimlik numarasının doğru girildiğinden emin olunuz.");
                }
            }
            else if (tcknTextBox.Text == "")
            {
                try
                {
                    aliciTcnVn = vergiDairesiTextBox.Text + "/" + vnTextBox.Text.Substring(0, 10);
                }
                catch (ArgumentOutOfRangeException)
                {
                    MessageBox.Show("Lütfen vergi numarasının doğru girildiğinden emin olunuz.");
                }
            }

            string aliciAdres = adresTextBox.Text;
            string aliciTel = telTextBox.Text;

            string sUnvan = "", sVDVN = "", sMersis = "", sAdres = "";
            if (saticiComboBox.Text == "Bereketli Topraklar")
            {
                sUnvan = "BEREKETLİ EMLAK SANAYİ VE TİCARET ANONİM ŞİRKETİ";
                sVDVN = "OSMANGAZİ/ 1650764758";
                sMersis = "0165076475800001";
                sAdres = "NİLÜFERKÖY MAH. NİLÜFER DEĞİRMEN SK. A BLOK NO: 12/7                                               Osmangazi/BURSA";
            }
            else if (saticiComboBox.Text == "BT Meta İnşaat")
            {
                sUnvan = "BT META İNŞAAT SANAYİ VE TİCARET ANONİM ŞİRKETİ";
                sVDVN = "KURTDERELİ/ 1871662733";
                sMersis = "0187166273300001";
                sAdres = "ÇAYIRHİSAR MAH. 6000 SK. NO: 6 İÇ KAPI NO: 3 ALTIEYLÜL/BALIKESİR";
            }
            else if (saticiComboBox.Text == "Bilal Aktaş")
            {
                sUnvan = "BİLAL AKTAŞ";
                sVDVN = "Nilüfer/ 2322798326";
                sMersis = "";
                sAdres = "İhsaniye mah. Leylak sk. No: 3B Daire: 13 Nilüfer/Bursa";
            }

            if (tcknTextBox.Text != "" && vnTextBox.Text != "")
            {
                MessageBox.Show("Aynı anda hem şirket hem de bireysel alıcı bilgisi girilemez!");
                return;
            }

            AliciBireysel aliciBireysel = new AliciBireysel();
            aliciBireysel = aliciBilgileriGetirFromComponents();

            AliciSirket aliciSirket = new AliciSirket();
            aliciSirket = aliciSirketBilgileriGetirFromComponents();


            if (BireyselAliciVarMi(tcknTextBox.Text) == false && tcknTextBox.Text != "" && tcknTextBox.Text.Length == 11 && vnTextBox.Text.Length < 10)
            {
                long tckn;
                if (long.TryParse(tcknTextBox.Text, out tckn))
                {
                    BireyselAliciEkle(aliciBireysel);
                    foreach (var item in aliciSirketComponentList)
                    {
                        if (item.Name != "adresTextBox" && item.Name != "telTextBox")
                        {
                            item.Text = "";
                        }
                    }
                }
                else
                {
                    MessageBox.Show("TC kimlik numarası yanlış girildi! Kişi bilgisi getirilemedi ve eklenemedi.");
                    return;
                }
            }
            else if (SirketAliciVarMi(vnTextBox.Text) == false && vnTextBox.Text != "" && vnTextBox.Text.Length == 10 && tcknTextBox.Text.Length < 11)
            {
                long vergiNumarasi;
                if (long.TryParse(vnTextBox.Text, out vergiNumarasi))
                {
                    AliciSirketEkle(aliciSirket);
                    foreach (var item in aliciBireyselComponentList)
                    {
                        if (item.Name != "adresTextBox" && item.Name != "telTextBox")
                        {
                            item.Text = "";
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Vergi numarası yanlış girildi! Şirket bilgisi getirilemedi ve eklenemedi.");
                    return;
                }
            }
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = null;

            doc = wordApp.Documents.Open(filePath);

            if (tasinmazSayisi % 2 == 0)
            {
                AddTablesBelowSentence(doc, "MADDE 2: SÖZLEŞMEYE KONU TAŞINMAZ", tasinmazSayisi / 2, 9, 4);
            }
            else if (tasinmazSayisi % 2 == 1)
            {
                AddTablesBelowSentence(doc, "MADDE 2: SÖZLEŞMEYE KONU TAŞINMAZ", (tasinmazSayisi + 1) / 2, 9, 4);
            }

            switch (ihtimal)
            {
                case 1:
                    pesinAyar(doc);
                    takasAyar(doc);
                    FindAndUpdateTaksitTable(taksitSayisi, doc);
                    break;
                case 2:
                    pesinAyar(doc);
                    takasAyar(doc);
                    break;
                case 3:
                    pesinAyar(doc);
                    FindAndUpdateTaksitTable(taksitSayisi, doc);
                    break;
                case 4:
                    takasAyar(doc);
                    FindAndUpdateTaksitTable(taksitSayisi, doc);
                    break;
                case 5:
                    pesinAyar(doc);
                    break;
                case 6:
                    takasAyar(doc);
                    break;
                case 7:
                    FindAndUpdateTaksitTable(taksitSayisi, doc);
                    break;
                default:
                    break;
            }


            // alıcı ve satıcı bilgilerini giren döngü
            foreach (ContentControl control in doc.ContentControls)
            {
                if (control.Tag == "sozlesmeTarihiText")
                {
                    control.Range.Text = sozlesmeTarihiDateTimePicker.Value.ToString("dd'/'MM'/'yyyy"); ;
                }
                if (control.Tag == "sUnvanText")
                {
                    control.Range.Text = sUnvan;
                }
                if (control.Tag == "sVDVNText")
                {
                    control.Range.Text = sVDVN;
                }
                if (control.Tag == "sMersisText")
                {
                    if (sMersis == "")
                    {
                        control.Range.Text = ".";
                        control.Range.Font.Size = 1;
                    }
                    else
                    {
                        control.Range.Text = sMersis;
                        control.Range.Font.Size = 11;
                    }
                }
                if (control.Tag == "sAdresText")
                {
                    control.Range.Text = sAdres;
                }

                if (control.Tag == "aAdSoyadUnvanText")
                {
                    control.Range.Text = aliciAdSoyadUnvan;
                }
                if (control.Tag == "aTCNVNText")
                {
                    control.Range.Text = aliciTcnVn;
                }
                if (control.Tag == "aAdresText")
                {
                    control.MultiLine = true; // İçerik denetimini çok satırlı yap
                    control.Range.Text = aliciAdres;
                }
                if (control.Tag == "aTelefonText")
                {
                    control.Range.Text = aliciTel;
                }

                if (control.Tag == "pesinCheckBox" && pesinC)
                {
                    control.Checked = true;
                }
                if (control.Tag == "takasCheckBox" && takasC)
                {
                    control.Checked = true;
                }
                if (control.Tag == "taksitliCheckBox" && taksitC)
                {
                    control.Checked = true;
                }
                if (control.Tag == "yaziylaParaMiktariText")
                {
                    control.Range.Text = NumberToCurrencyText(tasinmazBedeli);
                }
                if (control.Tag == "tasinmazSatisBedeliText")
                {
                    // Burada "N0" formatı, ondalık basamakları göstermez, sadece binlik ayırıcı uygular.
                    control.Range.Text = OndalikVarMi(tasinmazBedeli);
                }
            }

            // kodlarımı veri tabanı yapıma göre yeniden güncelle
            DateTime pesinOdemeTarihi = pesinOdemeTarihiDateTimePicker.Value;
            DateTime takasMalTeslimTarihi = takasMalTeslimTarihiDateTimePicker.Value;
            DateTime taksitBaslangicTarihi = taksitBaslangicTarihiDateTimePicker.Value;

            // taksitBaslangicTarihi.AddMonths(taksit sayısı)
            DateTime taksitSonlanmaTarihi = taksitBaslangicTarihi.AddMonths(taksitSayisi);
            string takasMalCinsi = takasMalCinsTextBox.Text;
            string takasMalOzellikleri = takasMalOzellikTextBox.Text;
            // taksit sayısı zaten var.


            try
            {
                // Geçersiz olan '/' yerine '-' kullanın
                string tarih = sozlesmeTarihiDateTimePicker.Value.ToString("dd-MM-yyyy");
                string fileName = aliciAdSoyadUnvan + " - " + tarih + " - Satış Sözleşmesi.docx";

                doc.SaveAs2(Path.Combine(kullanilanSozlesmelerYolu, fileName));
                doc.SaveAs2(Path.Combine(sozlesmeArsiviYolu, fileName));

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("Dosya kaydedilirken hata: Lütfen 'Kill All' ve 'Temizle' fonksiyonlarını kullandıktan sonra yeniden deneyiniz.");
                return;
            }

            Sozlesme sozlesme = new Sozlesme();

            switch (ihtimal)
            {
                //pesin takas taksit
                case 1:
                    sozlesme = sozlesmeDegiskeni(true, true, true);
                    break;
                //pesin takas
                case 2:
                    sozlesme = sozlesmeDegiskeni(true, true, false);
                    break;
                // pesin taksit 
                case 3:
                    sozlesme = sozlesmeDegiskeni(true, false, true);
                    break;
                case 4:
                    sozlesme = sozlesmeDegiskeni(false, true, true);
                    break;
                case 5:
                    sozlesme = sozlesmeDegiskeni(true, false, false);
                    break;
                case 6:
                    sozlesme = sozlesmeDegiskeni(false, true, false);
                    break;
                case 7:
                    sozlesme = sozlesmeDegiskeni(false, false, true);
                    break;
                default:
                    break;
            }

            if (tcknTextBox.Text != "")
            {
                sozlesme.AliciTc = tcknTextBox.Text.Substring(0, 11);
            }
            else if (vnTextBox.Text != "")
            {
                sozlesme.AliciVN = vnTextBox.Text.Substring(0, 10);
            }
            sozlesme.SozlesmeyiOlusturanKullanici = GirisForm.kullanici;
            sozlesme.TasinmazBedeli = tasinmazBedeli;
            sozlesme.SaticiUnvan = saticiComboBox.Text;
            SozlesmeEkle(sozlesme);

            // tarih ekle

            try
            {
                doc?.Close(ref missing, ref missing, ref missing);
                wordApp.Application.Quit(ref missing, ref missing, ref missing);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);

                await System.Threading.Tasks.Task.Delay(1000);
                notifyIcon1.Icon = SystemIcons.Information; // Veya kendi .ico dosyanız
                notifyIcon1.Visible = true;
                notifyIcon1.Text = "Sözleşme Üretimi";
                notifyIcon1.BalloonTipTitle = "Belge Üretimi";
                notifyIcon1.BalloonTipText = "Belge üretimi başarıyla tamamlandı.";
                notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
                notifyIcon1.ShowBalloonTip(5000);
                await System.Threading.Tasks.Task.Delay(5500);
                notifyIcon1.Visible = false;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("İşlem başarısız, lütfen 'Temizle' ve 'Kill All' butonlarını kullandıktan sonra tekrar deneyiniz.");
                throw;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            //wordApp.Quit();
        }
        string OndalikVarMi(decimal deger)
        {
            // Bir sayının tam kısımdan arta kalan ondalık kısmı 0'dan farklı mı?
            bool x = false;
            if ((deger % 1) != 0)
            {
                x = true;
            }
            CultureInfo turkishCulture = new CultureInfo("tr-TR");

            string formatliDeger;

            if (x)
            {
                formatliDeger = deger.ToString("N2", turkishCulture);
            }
            else
            {
                formatliDeger = deger.ToString("N0", turkishCulture);
            }

            return formatliDeger;
        }
        bool IsFileLocked(string filePath)
        {
            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    stream.Close();
                }
                return false;
            }
            catch (IOException ex)
            {
                MessageBox.Show("Dosyanın açık olup olmadığını kontrol ederken hata. Hata ayrıntıları: " + ex.StackTrace);
                return true;
            }
        }


        private static string masaustuYolu = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        // "Kullanılan Sözleşmeler" klasörünün tam yolu
        string kullanilanSozlesmelerYolu = Path.Combine(masaustuYolu, "Kullanılan Sözleşmeler");

        // "Yedek Sözleşmeler" klasörünün tam yolu
        string sozlesmeArsiviYolu = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sözleşme Arşivi");
        public void klasorOlustur()
        {
            // Kullanıcının masaüstü dizinini al

            // "Kullanılan Sözleşmeler" klasörünü oluştur
            if (!Directory.Exists(kullanilanSozlesmelerYolu))
            {
                Directory.CreateDirectory(kullanilanSozlesmelerYolu);
            }

            // "Yedek Sözleşmeler" klasörünü oluştur
            if (!Directory.Exists(sozlesmeArsiviYolu))
            {
                Directory.CreateDirectory(sozlesmeArsiviYolu);
            }
        }


        public Sozlesme sozlesmeDegiskeni(bool pesin, bool takas, bool taksit)
        {
            int taksitSayisi = 0;
            if (int.TryParse(taksitSayisiComboBox.Text, out taksitSayisi))
            {

            }

            decimal taksitTutari;
            if (decimal.TryParse(taksitTutariTextBox.Text, out taksitTutari))
            {

            }

            decimal takasMalTutari;
            if (decimal.TryParse(takasMalTutarTextBox.Text, out takasMalTutari))
            {

            }
            decimal pesinOdemeTutari;
            if (decimal.TryParse(pesinOdemeTutariTextBox.Text, out pesinOdemeTutari))
            {

            }



            DateTime pesinOdemeTarihi = pesinOdemeTarihiDateTimePicker.Value;
            DateTime takasMalTeslimTarihi = takasMalTeslimTarihiDateTimePicker.Value;
            DateTime taksitBaslangicTarihi = taksitBaslangicTarihiDateTimePicker.Value;
            DateTime sozlesmeTarihi = sozlesmeTarihiDateTimePicker.Value;
            // son componenti
            // veya
            // taksitBaslangicTarihi.AddMonths(taksit sayısı)
            DateTime taksitSonlanmaTarihi = taksitBaslangicTarihi.AddMonths(taksitSayisi);
            string takasMalCinsi = takasMalCinsTextBox.Text;
            string takasMalOzellikleri = takasMalOzellikTextBox.Text;

            Sozlesme sozlesme = new Sozlesme();

            if (pesin)
            {
                sozlesme.PesinOdemeTutari = pesinOdemeTutari;
                sozlesme.PesinOdemeTarihi = pesinOdemeTarihi;
            }

            if (takas)
            {
                sozlesme.TakasMalinCinsi = takasMalCinsi;
                sozlesme.TakasMalinOzellikleri = takasMalOzellikleri;
                sozlesme.TakasMalinTutari = takasMalTutari;
                sozlesme.TakasMalinTeslimTarihi = takasMalTeslimTarihi;
            }

            if (taksit)
            {
                sozlesme.TaksitSayisi = byte.Parse(taksitSayisi.ToString());
                sozlesme.TaksitTutari = taksitTutari;
                sozlesme.TaksitBaslangicTarihi = taksitBaslangicTarihi;
                sozlesme.TaksitBitisTarihi = taksitSonlanmaTarihi;
            }
            //sozlesme.TasinmazBedeli = tasinmazBedeli;
            sozlesme.SozlesmeTarihi = sozlesmeTarihi;

            return sozlesme;
        }

        private void AddTablesBelowSentence(Document wordDoc, string sentence, int numberOfTables, int rows, int columns)
        {
            // Cümleyi bulmak için Range oluştur
            Range searchRange = wordDoc.Content;

            if (searchRange.Find.Execute(sentence))
            {
                // Cümleyi bulduktan sonra aralığın bitişinden itibaren tablo ekleme noktasını belirle
                Range insertionPoint = searchRange.Duplicate;
                insertionPoint.Collapse(WdCollapseDirection.wdCollapseEnd);
                string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                int count = 0;
                int tasinmazSayisi = int.Parse(tasinmazSayisiComboBox.Text);
                int countXoldugunda = 0;
                if (tasinmazSayisi % 2 == 1)
                {
                    countXoldugunda = tasinmazSayisi - 1;
                }
                int countSutun1 = 0;
                int countSutun2 = 1;
                for (int t = 0; t < numberOfTables; t++)
                {
                    List<Control> tasinmazTextBoxList = new List<Control>();
                    //foreach (var (id, control) in values)
                    //{
                    //    if (id == t)
                    //    {
                    //        tasinmazTextBoxList.Add(control);
                    //    }
                    //}
                    // Tabloyu ekle
                    Table newTable = wordDoc.Tables.Add(insertionPoint, rows, columns);
                    newTable.Columns[1].PreferredWidth = wordDoc.Application.CentimetersToPoints(2.9f);
                    newTable.Columns[2].PreferredWidth = wordDoc.Application.CentimetersToPoints(5.53f);
                    newTable.Columns[3].PreferredWidth = wordDoc.Application.CentimetersToPoints(2.9f);
                    newTable.Columns[4].PreferredWidth = wordDoc.Application.CentimetersToPoints(5.53f);
                    // Tablo biçimlendirme
                    newTable.Borders.Enable = 0; // Kenarlıkları etkinleştir
                    newTable.Rows.Alignment = WdRowAlignment.wdAlignRowLeft; // Sola hizala
                    newTable.LeftPadding = 0; // Sol kenar boşluğu sıfırla
                    newTable.AllowPageBreaks = false; // Sayfa taşmasını engelle

                    // Tablo genişliğini sayfa genişliğine göre ayarla
                    float pageWidth = wordDoc.PageSetup.PageWidth - wordDoc.PageSetup.LeftMargin - wordDoc.PageSetup.RightMargin;
                    newTable.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
                    newTable.PreferredWidth = pageWidth;
                    // Tablo içeriğini doldur

                    for (int i = 1; i <= rows; i++)
                    {
                        for (int j = 1; j <= columns; j++)
                        {
                            // tasinmaz bilgileri 2. sütun.
                            if (j == 2)
                            {
                                //MessageBox.Show(textBoxList[countSutun1][0].ToString());
                                //foreach (var item in textBoxList[countSutun1])
                                //{
                                //    MessageBox.Show(item.Text);
                                //}
                                switch (i)// taşınmaz sayısının tek olmasına bağlı olarak bir güncelleme yapacağız
                                {
                                    case 2:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun1][i - 2].Text;
                                        break;
                                    case 3:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun1][i - 2].Text;
                                        break;
                                    case 4:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun1][i - 2].Text;
                                        break;
                                    case 5:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun1][i - 2].Text;
                                        break;
                                    case 6:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun1][i - 2].Text;
                                        break;
                                    case 7:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun1][i - 2].Text;
                                        break;
                                    case 8:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun1][i - 2].Text;
                                        break;
                                    case 9:
                                        decimal tasinmazFiyat;
                                        string formatliDeger = "";
                                        if (decimal.TryParse(textBoxList[countSutun1][i - 2].Text, out tasinmazFiyat))
                                        {
                                            formatliDeger = OndalikVarMi(decimal.Parse(textBoxList[countSutun1][i - 2].Text));
                                        }
                                        else
                                        {
                                            formatliDeger = "";
                                        }
                                        newTable.Cell(i, j).Range.Text = formatliDeger + "-TL";

                                        //newTable.Cell(i, j).Range.Text = textBoxList[countSutun1][i - 2].Text + "-TL";
                                        countSutun1 += 2;
                                        count++;
                                        break;
                                }
                            }

                            // tasinmaz bilgileri 4. sütun. eğer tasinmaz sayısı tek sayı ise
                            if ((j == 4 && count + 1 < tasinmazSayisi) && tasinmazSayisi % 2 == 1 && tasinmazSayisi != 1)
                            {
                                switch (i)// taşınmaz sayısının tek olmasına bağlı olarak bir güncelleme yapacağız
                                {
                                    case 2:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text;
                                        break;
                                    case 3:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text;
                                        break;
                                    case 4:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text;
                                        break;
                                    case 5:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text;
                                        break;
                                    case 6:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text;
                                        break;
                                    case 7:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text;
                                        break;
                                    case 8:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text;
                                        break;
                                    case 9:
                                        decimal tasinmazFiyat;
                                        string formatliDeger = "";
                                        if (decimal.TryParse(textBoxList[countSutun2][i - 2].Text, out tasinmazFiyat))
                                        {
                                            formatliDeger = OndalikVarMi(decimal.Parse(textBoxList[countSutun2][i - 2].Text));
                                        }
                                        else
                                        {
                                            formatliDeger = "";
                                        }

                                        newTable.Cell(i, j).Range.Text = formatliDeger + "-TL";

                                        //newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text + "-TL";
                                        countSutun2 += 2;
                                        count++;
                                        break;
                                }
                            }

                            // tasinmaz bilgileri 4. sütun. eğer tasinmaz sayısı çift sayı ise
                            if (j == 4 && tasinmazSayisi % 2 == 0)
                            {
                                switch (i)// taşınmaz sayısının tek olmasına bağlı olarak bir güncelleme yapacağız
                                {
                                    case 2:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text;
                                        break;
                                    case 3:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text;
                                        break;
                                    case 4:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text;
                                        break;
                                    case 5:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text;
                                        break;
                                    case 6:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text;
                                        break;
                                    case 7:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text;
                                        break;
                                    case 8:
                                        newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text;
                                        break;
                                    case 9:
                                        decimal tasinmazFiyat;
                                        string formatliDeger = "";
                                        if (decimal.TryParse(textBoxList[countSutun2][i - 2].Text, out tasinmazFiyat))
                                        {
                                            formatliDeger = OndalikVarMi(decimal.Parse(textBoxList[countSutun2][i - 2].Text));
                                        }
                                        newTable.Cell(i, j).Range.Text = formatliDeger + "-TL";

                                        //newTable.Cell(i, j).Range.Text = textBoxList[countSutun2][i - 2].Text + "-TL";
                                        countSutun2 += 2;
                                        count++;
                                        break;
                                }
                            }



                            // il ilçe 1. sütun
                            if (j == 1)
                            {
                                switch (i)// taşınmaz sayısının tek olmasına bağlı olarak bir güncelleme yapacağız
                                {
                                    case 1:
                                        if (j == 1)
                                        {
                                            newTable.Cell(i, j).Range.Text = alphabet[count] + ")";
                                        }
                                        else // j == 3
                                        {
                                            newTable.Cell(i, j).Range.Text = alphabet[count + 1] + ")";
                                        }
                                        break;
                                    case 2:
                                        newTable.Cell(i, j).Range.Text = "İl                                :";
                                        break;
                                    case 3:
                                        newTable.Cell(i, j).Range.Text = "İlçe                             :";
                                        break;
                                    case 4:
                                        newTable.Cell(i, j).Range.Text = "Mahalle                   :";
                                        break;
                                    case 5:
                                        newTable.Cell(i, j).Range.Text = "Ada                            :";
                                        break;
                                    case 6:
                                        newTable.Cell(i, j).Range.Text = "Parsel                        :";
                                        break;
                                    case 7:
                                        newTable.Cell(i, j).Range.Text = "Cinsi                        :";
                                        break;
                                    case 8:
                                        newTable.Cell(i, j).Range.Text = "Metrekare              :";
                                        break;
                                    case 9:
                                        newTable.Cell(i, j).Range.Text = "Fiyat                         :";
                                        //count++;
                                        break;
                                }
                            }


                            // il ilçe 3. sütun. eğer tasinmaz sayısı tek sayı ise
                            if ((j == 3 && count + 1 < tasinmazSayisi) && tasinmazSayisi % 2 == 1)
                            {
                                switch (i)// taşınmaz sayısının tek olmasına bağlı olarak bir güncelleme yapacağız
                                {
                                    case 1:
                                        if (j == 1)
                                        {
                                            newTable.Cell(i, j).Range.Text = alphabet[count] + ")";
                                        }
                                        else // j == 3
                                        {
                                            newTable.Cell(i, j).Range.Text = alphabet[count + 1] + ")";
                                        }
                                        break;
                                    case 2:
                                        newTable.Cell(i, j).Range.Text = "İl                                :";
                                        break;
                                    case 3:
                                        newTable.Cell(i, j).Range.Text = "İlçe                             :";
                                        break;
                                    case 4:
                                        newTable.Cell(i, j).Range.Text = "Mahalle                   :";
                                        break;
                                    case 5:
                                        newTable.Cell(i, j).Range.Text = "Ada                            :";
                                        break;
                                    case 6:
                                        newTable.Cell(i, j).Range.Text = "Parsel                        :";
                                        break;
                                    case 7:
                                        newTable.Cell(i, j).Range.Text = "Cinsi                        :";
                                        break;
                                    case 8:
                                        newTable.Cell(i, j).Range.Text = "Metrekare              :";
                                        break;
                                    case 9:
                                        newTable.Cell(i, j).Range.Text = "Fiyat                         :";
                                        //count++;
                                        break;
                                }
                            }


                            // il ilçe 3. sütun. eğer tasinmaz sayısı çift sayı ise
                            if (j == 3 && tasinmazSayisi % 2 == 0)
                            {
                                switch (i)// taşınmaz sayısının tek olmasına bağlı olarak bir güncelleme yapacağız
                                {
                                    case 1:
                                        if (j == 1)
                                        {
                                            newTable.Cell(i, j).Range.Text = alphabet[count] + ")";
                                        }
                                        else // j == 3
                                        {
                                            newTable.Cell(i, j).Range.Text = alphabet[count + 1] + ")";
                                        }
                                        break;
                                    case 2:
                                        newTable.Cell(i, j).Range.Text = "İl                                :";
                                        break;
                                    case 3:
                                        newTable.Cell(i, j).Range.Text = "İlçe                             :";
                                        break;
                                    case 4:
                                        newTable.Cell(i, j).Range.Text = "Mahalle                   :";
                                        break;
                                    case 5:
                                        newTable.Cell(i, j).Range.Text = "Ada                            :";
                                        break;
                                    case 6:
                                        newTable.Cell(i, j).Range.Text = "Parsel                        :";
                                        break;
                                    case 7:
                                        newTable.Cell(i, j).Range.Text = "Cinsi                        :";
                                        break;
                                    case 8:
                                        newTable.Cell(i, j).Range.Text = "Metrekare              :";
                                        break;
                                    case 9:
                                        newTable.Cell(i, j).Range.Text = "Fiyat                         :";
                                        //count++;
                                        break;
                                }
                            }

                        }
                    }

                    // Tabloyu taşma olmaması için birleştir
                    foreach (Row row in newTable.Rows)
                    {
                        row.Range.ParagraphFormat.KeepWithNext = -1; // Bir sonrakiyle birlikte tut
                        row.Range.ParagraphFormat.KeepTogether = -1; // Satırı bölme
                    }

                    // Yeni tabloyu ekledikten sonra sonraki ekleme noktasını güncelle
                    insertionPoint = newTable.Range;
                    insertionPoint.Collapse(WdCollapseDirection.wdCollapseEnd);

                    // Tablolar arasında boşluk oluştur
                    insertionPoint.InsertParagraphAfter();
                    insertionPoint.Collapse(WdCollapseDirection.wdCollapseEnd);
                }
            }
            else
            {
                MessageBox.Show("Belirtilen cümle bulunamadı.");
            }
        }


        private void pesinAyar(Document dokumanAyar)
        {
            foreach (ContentControl control in dokumanAyar.ContentControls)
            {
                if (control.Tag == "pesinSatisOdemeTutari")
                {
                    decimal pesinTutar;
                    if (decimal.TryParse(pesinOdemeTutariTextBox.Text, out pesinTutar))
                    {
                        control.Range.Text = OndalikVarMi(pesinTutar) + "-TL (" + NumberToCurrencyText(pesinTutar).ToString();
                        control.Range.Font.Bold = 1;
                    }
                    else
                    {
                        MessageBox.Show("Peşin tutarı doğru biçimde girilmedi!");
                        return;
                    }
                }
                if (control.Tag == "pesinSatisOdemeTarihi")
                {
                    control.Range.Text = pesinOdemeTarihiDateTimePicker.Value.ToString("dd'/'MM'/'yyyy");
                    control.Range.Font.Bold = 1;
                }
            }

            tabloBirlesikKalsin(dokumanAyar, "PEŞİN SATIŞ");
        }
        private void takasAyar(Document dokumanAyar)
        {
            foreach (ContentControl control in dokumanAyar.ContentControls)
            {
                if (control.Tag == "takasMalinCinsiText")
                {
                    control.MultiLine = true; // İçerik denetimini çok satırlı yap
                    control.Range.Text = takasMalCinsTextBox.Text;
                    control.Range.Font.Bold = 1;
                }
                if (control.Tag == "takasMalinOzellikleriText")
                {
                    control.MultiLine = true; // İçerik denetimini çok satırlı yap
                    control.Range.Text = takasMalOzellikTextBox.Text;
                    control.Range.Font.Bold = 1;
                }
                if (control.Tag == "takasMalinTutariText")
                {
                    //control.MultiLine = true; // İçerik denetimini çok satırlı yap
                    decimal takasMalTutari;
                    if (decimal.TryParse(takasMalTutarTextBox.Text, out takasMalTutari))
                    {
                        control.Range.Text = OndalikVarMi(takasMalTutari) + "-TL (" + NumberToCurrencyText(takasMalTutari);
                        control.Range.Font.Bold = 1;
                    }
                    else
                    {
                        MessageBox.Show("Takas mal tutarı doğru biçimde girilmedi!");
                        return;
                    }
                }
                if (control.Tag == "takasMalinTeslimTarihiText")
                {
                    control.Range.Text = takasMalTeslimTarihiDateTimePicker.Value.ToString("dd'/'MM'/'yyyy");
                    control.Range.Font.Bold = 1;
                }
            }
            tabloBirlesikKalsin(dokumanAyar, "TAKAS SATIŞI");
        }

        private void tabloBirlesikKalsin(Document wordDoc, string tabloBaslik)
        {
            Table taksitTable = null;

            // Belgedeki tüm tabloları kontrol et
            foreach (Table table in wordDoc.Tables)
            {
                // İlk satırdaki hücrede "Taksit No" olup olmadığını kontrol et
                if (table.Cell(1, 1).Range.Text.Trim().Contains(tabloBaslik))
                {
                    taksitTable = table;
                    break; // İlgili tabloyu bulduk, çıkalım
                }
            }

            if (taksitTable != null)
            {
                // Tabloyu güncelle
                //UpdateTaksitTable(taksitTable, taksitSayisi, wordDoc);

                foreach (Row row in taksitTable.Rows)
                {
                    row.Range.ParagraphFormat.KeepWithNext = -1; // Bir sonrakiyle birlikte tut
                    row.Range.ParagraphFormat.KeepTogether = -1; // Satırı bölme
                }
            }
        }


        private void FindAndUpdateTaksitTable(int taksitSayisi, Document wordDoc)
        {
            Table taksitTable = null;

            // Belgedeki tüm tabloları kontrol et
            foreach (Table table in wordDoc.Tables)
            {
                // İlk satırdaki hücrede "Taksit No" olup olmadığını kontrol et
                if (table.Cell(1, 1).Range.Text.Trim().Contains("TAKSİTLİ SATIŞ"))
                {
                    taksitTable = table;
                    break; // İlgili tabloyu bulduk, çıkalım
                }
            }

            if (taksitTable != null)
            {
                // Tabloyu güncelle
                UpdateTaksitTable(taksitTable, taksitSayisi, wordDoc);
            }
            else
            {
                MessageBox.Show("Tablo bulunamadı.");
            }
        }


        private void UpdateTaksitTable(Table taksitTable, int taksitSayisi, Document wordDoc)
        {
            // Önce tablonun tüm kenarlıklarını sıfırla (No Border)
            taksitTable.Borders.Enable = 1;

            // Tabloya çerçeve ekle (Dış kenarlıklar)
            taksitTable.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
            taksitTable.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
            taksitTable.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
            taksitTable.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
            taksitTable.Borders[WdBorderType.wdBorderVertical].LineStyle = WdLineStyle.wdLineStyleNone; // İç dikey kenarlıklar
            taksitTable.Borders[WdBorderType.wdBorderHorizontal].LineStyle = WdLineStyle.wdLineStyleNone; // İç yatay kenarlıklar

            // Taksit sayısına göre tabloyu genişlet
            while (taksitTable.Rows.Count < taksitSayisi + 1) // Başlık dahil
            {
                taksitTable.Rows.Add();
            }
            taksitTable.Rows.Add();

            taksitTable.Rows.Alignment = WdRowAlignment.wdAlignRowLeft; // Tablo sola hizalanır
            taksitTable.LeftPadding = 0; // Sol kenar boşluğu sıfırlanır

            float pageWidth = wordDoc.PageSetup.PageWidth - wordDoc.PageSetup.LeftMargin - wordDoc.PageSetup.RightMargin;
            taksitTable.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
            taksitTable.PreferredWidth = pageWidth;

            for (int i = 1; i <= 3; i++)
            {
                taksitTable.Cell(2, i).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                taksitTable.Cell(2, i).Range.Font.Name = "Times New Roman";
                taksitTable.Cell(2, i).Range.Font.Bold = 1;
            }

            taksitTable.Cell(1, 1).Range.Text = "TAKSİTLİ SATIŞ";
            taksitTable.Cell(2, 2).Range.Text = "TUTAR";
            taksitTable.Cell(2, 3).Range.Text = "TARİH";

            taksitTable.Cell(1, 1).Range.Font.Name = "Times New Roman";
            taksitTable.Cell(1, 1).Range.Font.Bold = 1;

            //taksitTable.Columns[1].PreferredWidth = wordDoc.Application.CentimetersToPoints(4); // Taksit No
            //taksitTable.Columns[2].PreferredWidth = wordDoc.Application.CentimetersToPoints(7); // Taksit Tutarı
            //taksitTable.Columns[3].PreferredWidth = wordDoc.Application.CentimetersToPoints(10);

            taksitTable.Cell(1, 1).Merge(taksitTable.Cell(1, 3));

            DateTime taksitBaslangicTarihi = taksitBaslangicTarihiDateTimePicker.Value;

            // Her satır ve sütun için veri gir
            for (int i = 0; i < taksitSayisi; i++)
            {
                if (i < 9)
                {
                    taksitTable.Cell(i + 3, 1).Range.Text = "     " + (i + 1).ToString() + ". Taksit  : "; // Taksit No
                }
                else
                {
                    taksitTable.Cell(i + 3, 1).Range.Text = "   " + (i + 1).ToString() + ". Taksit  : "; // Taksit No
                }


                if (taksitTutariComponentList.Count > 0)
                {
                    taksitTable.Cell(i + 3, 2).Range.Text = "   " + taksitTutariComponentList[i].Text + "-TL";             // Taksit Tutarı
                }
                else
                {
                    taksitTable.Cell(i + 3, 2).Range.Text = "  Girilmedi. ";             // Taksit Tutarı
                }


                taksitTable.Cell(i + 3, 3).Range.Text = dtpList[i].Value.ToString("dd'/'MM'/'yyyy"); ; // Taksit Tarihi


                // Metin hizalamalarını ayarla
                taksitTable.Cell(i + 3, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                taksitTable.Cell(i + 3, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                taksitTable.Cell(i + 3, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                taksitTable.Cell(i + 3, 1).Range.Font.Bold = 1;
                taksitTable.Cell(i + 3, 2).Range.Font.Bold = 1;
                taksitTable.Cell(i + 3, 3).Range.Font.Bold = 1;

                taksitTable.Cell(i + 3, 1).Range.Font.Name = "Times New Roman";
                taksitTable.Cell(i + 3, 2).Range.Font.Name = "Times New Roman";
                taksitTable.Cell(i + 3, 3).Range.Font.Name = "Times New Roman";
            }

            foreach (Row row in taksitTable.Rows)
            {
                row.Range.ParagraphFormat.KeepWithNext = -1; // Bir sonrakiyle birlikte tut
                row.Range.ParagraphFormat.KeepTogether = -1; // Satırı bölme
            }
            // Tablo hizalaması ve genel çerçeve ayarları
            taksitTable.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
            taksitTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            taksitTable.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth150pt;

            //taksitTable.Borders.Enable = 0;
        }
        public void taksitControls()
        {
            // Yeni kontroller eklenmeden önce mevcutları temizle
            fPnl.Controls.Clear();
            lblList.Clear();
            txtList.Clear();
            dtpList.Clear();

            DateTime startDate = taksitBaslangicTarihiDateTimePicker.Value;

            Random rnd = new Random();
            Label lbl;

            int taksitSayisi;
            if (int.TryParse(taksitSayisiComboBox.Text, out taksitSayisi))
            {
                for (int i = 0; i < taksitSayisi; i++)
                {
                    int yOffset = 90; // Y ekseni için başlangıç yüksekliği
                    int ySpacing = 30; // Her kontrol grubu arasındaki boşluk

                    if (i < 9)
                    {
                        lbl = new Label
                        {
                            Text = $"{i + 1}. Taksit Tutarı: ",
                            Padding = new Padding(7, 13, 0, 0),
                            AutoSize = true,
                            Location = new System.Drawing.Point(0, yOffset + (i * ySpacing)) // Y ekseni için hesaplama
                        };
                    }
                    else
                    {
                        lbl = new Label
                        {
                            Text = $"{i + 1}. Taksit Tutarı: ",
                            Padding = new Padding(0, 13, 0, 0),
                            AutoSize = true,
                            Location = new System.Drawing.Point(10, yOffset + (i * ySpacing)) // Y ekseni için hesaplama
                        };
                    }


                    TextBox txt = new TextBox
                    {
                        AutoSize = true,
                        Margin = new Padding(0, 10, 0, 0),
                        Width = 100,
                        Location = new System.Drawing.Point(200, yOffset + (i * ySpacing)) // Aynı Y ekseni düzeni
                    };

                    DateTimePicker dtp = new DateTimePicker
                    {
                        AutoSize = true,
                        Value = startDate.AddMonths(i),
                        Format = DateTimePickerFormat.Short,
                        Margin = new Padding(15, 10, 0, 0),
                        Width = 150,
                        Location = new System.Drawing.Point(320, yOffset + (i * ySpacing)) // Aynı Y ekseni düzeni
                    };

                    // Listelere ve panele ekle
                    lblList.Add(lbl);
                    txtList.Add(txt);
                    dtpList.Add(dtp);


                    fPnl.Controls.Add(lbl);
                    fPnl.Controls.Add(txt);
                    fPnl.Controls.Add(dtp);

                    foreach (var item in fPnl.Controls)
                    {
                        if (item is TextBox txt2)
                        {
                            if (taksitTutariTextBox.Text != "")
                            {
                                txt2.Text = taksitTutariTextBox.Text;
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Geçerli bir sayı giriniz.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }


        private void pesinCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (pesinCheckBox.Checked)
            {
                ActivateControls(pesinComponentList);
            }
            else
            {
                DeactivateControls(pesinComponentList);
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            tcknTextBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            tcknTextBox.AutoCompleteSource = AutoCompleteSource.CustomSource;

            vnTextBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            vnTextBox.AutoCompleteSource = AutoCompleteSource.CustomSource;

            aliciBireyselComponentList.Add(adSoyadTextBox);
            aliciBireyselComponentList.Add(tcknTextBox);
            aliciBireyselComponentList.Add(telTextBox);
            aliciBireyselComponentList.Add(adresTextBox);

            aliciSirketComponentList.Add(unvanTextBox);
            aliciSirketComponentList.Add(vergiDairesiTextBox);
            aliciSirketComponentList.Add(vnTextBox);
            aliciSirketComponentList.Add(telTextBox);
            aliciSirketComponentList.Add(adresTextBox);

            foreach (var item in this.Controls)
            {
                if (item is DateTimePicker dateTimePicker)
                {
                    // Format ve CustomFormat ayarlarını yapıyoruz
                    dateTimePicker.Format = DateTimePickerFormat.Short;
                    //dateTimePicker.CustomFormat = "dd/MM/yyyy"; // Tarih formatı
                }
            }
            DeactivateAllControls();
        }

        public void DeactivateAllControls()
        {
            foreach (var item in pesinComponentList)
            {
                item.Enabled = false;
            }

            foreach (var item in takasComponentList)
            {
                item.Enabled = false;
            }

            foreach (var item in taksitComponentList)
            {
                item.Enabled = false;
            }
        }
        public void ActivateControls(List<Control> controls)
        {
            foreach (var item in controls)
            {
                item.Enabled = true;
            }
        }
        public void DeactivateControls(List<Control> controls)
        {
            foreach (var item in controls)
            {
                item.Enabled = false;
            }
        }

        private void tasinmazSayisiComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComponentUret();
        }

        public void ComponentUret()
        {
            fpnl2.Controls.Clear();
            textBoxList = new List<List<TextBox>>();
            Random rnd = new Random();
            Label lbl;

            Padding pad = new Padding(0, 0, 0, 0);

            int tasinmazSayisi;
            string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            int count = 0;
            if (int.TryParse(tasinmazSayisiComboBox.Text, out tasinmazSayisi))
            {
                for (int j = 0; j < tasinmazSayisi; j++)
                {
                    List<TextBox> matrix = new List<TextBox>();
                    for (int i = 0; i < 9; i++)
                    {
                        int yOffset = 90; // Y ekseni için başlangıç yüksekliği
                        int ySpacing = 30; // Her kontrol grubu arasındaki boşluk

                        string labelText = "";

                        switch (i)
                        {
                            case 0:
                                labelText = alphabet[count].ToString() + ")";
                                count++;
                                break;
                            case 1:
                                labelText = "İl";
                                break;
                            case 2:
                                labelText = "İlçe";
                                break;
                            case 3:
                                labelText = "Mahalle";
                                break;
                            case 4:
                                labelText = "Ada";
                                break;
                            case 5:
                                labelText = "Parsel";
                                break;
                            case 6:
                                labelText = "Cinsi";
                                break;
                            case 7:
                                labelText = "Metrekare";
                                break;
                            case 8:
                                labelText = "Fiyat";
                                break;
                            default:
                                break;
                        }

                        TextBox txt;


                        if (i == 8)
                        {
                            lbl = new Label
                            {
                                Text = labelText,
                                //Margin = new Padding(0, 10, 0, 0),
                                AutoSize = false,
                                Margin = new Padding(0, 0, 0, 20),
                                Location = new System.Drawing.Point(0, yOffset + (i * ySpacing)) // Y ekseni için hesaplama
                            };

                            txt = new TextBox
                            {
                                AutoSize = false,
                                Margin = new Padding(0, 0, 0, 20),
                                Width = 100,
                                Location = new System.Drawing.Point(200, yOffset + (i * ySpacing)) // Aynı Y ekseni düzeni
                            };
                        }
                        else
                        {
                            lbl = new Label
                            {
                                Text = labelText,
                                Margin = new Padding(0, 0, 0, 0),
                                AutoSize = false,
                                Location = new System.Drawing.Point(0, yOffset + (i * ySpacing)) // Y ekseni için hesaplama
                            };
                            bool z = true;
                            if (i == 0)
                            {
                                z = false;
                            }
                            txt = new TextBox
                            {
                                AutoSize = false,
                                Margin = new Padding(0, 0, 0, 0),
                                Width = 100,
                                Enabled = z,
                                //Visible = z,
                                Location = new System.Drawing.Point(200, yOffset + (i * ySpacing)) // Aynı Y ekseni düzeni
                            };
                        }

                        if (i != 0)
                        {
                            matrix.Add(txt);
                            txtTasinmazList.Add(txt);
                        }

                        fpnl2.Controls.Add(lbl);
                        fpnl2.Controls.Add(txt);
                    }
                    textBoxList.Add(matrix);
                }
            }
            else
            {
                MessageBox.Show("Geçerli bir sayı giriniz.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DynamicTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                decimal taksitToplam = 0;
                foreach (var item in taksitTutariComponentList)
                {
                    taksitToplam += decimal.Parse(item.Text);
                }
                taksitToplamLabel.Text = "Toplam Taksit Tutarı: " + taksitToplam.ToString() + " TL";
            }
            catch (Exception)
            {
                taksitToplamLabel.Text = "Toplam Taksit Tutarı: Hesaplanamadı";
            }
        }
        private void taksitCheckBox_CheckedChanged_1(object sender, EventArgs e)
        {
            if (taksitCheckBox.Checked)
            {
                ActivateControls(taksitComponentList);
            }
            else
            {
                DeactivateControls(taksitComponentList);
            }
        }

        private void taksitSayisiComboBox_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            taksitControls();

            taksitTutariComponentList.Clear();
            foreach (var item in fPnl.Controls)
            {
                if (item is TextBox textbox)
                {
                    // Format ve CustomFormat ayarlarını yapıyoruz
                    taksitTutariComponentList.Add(textbox);
                    textbox.KeyUp += DynamicTextBox_KeyUp;
                }
            }

            try
            {
                decimal taksitToplam = 0;
                foreach (var item in taksitTutariComponentList)
                {
                    taksitToplam += decimal.Parse(item.Text);
                }
                taksitToplamLabel.Text = "Toplam Taksit Tutarı: " + taksitToplam.ToString() + " TL";
            }
            catch (Exception)
            {
                taksitToplamLabel.Text = "Toplam Taksit Tutarı: Hesaplanamadı";
            }
        }

        private void takasCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (takasCheckBox.Checked)
            {
                ActivateControls(takasComponentList);
            }
            else
            {
                DeactivateControls(takasComponentList);
            }
        }


        private void tcknTextBox_Enter(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(connectionString);

            con.Open();

            SqlCommand cmd = new SqlCommand("select aliciTc, adSoyad from alici_bireysel where aliciTc like '" + tcknTextBox.Text + "%'", con);

            SqlDataReader dr = cmd.ExecuteReader();

            AutoCompleteStringCollection scollection = new AutoCompleteStringCollection();

            while (dr.Read())
            {
                string x = dr["aliciTc"].ToString();
                string y = dr["adSoyad"].ToString();
                scollection.Add(x + " " + y);
            }

            tcknTextBox.AutoCompleteCustomSource = scollection;

            con.Close();
        }

        private void vnTextBox_Enter(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(connectionString);

            con.Open();

            SqlCommand cmd = new SqlCommand("select vergiNumarasi, sirketUnvan from alici_sirket where vergiNumarasi like '" + vnTextBox.Text + "%'", con);

            SqlDataReader dr = cmd.ExecuteReader();

            AutoCompleteStringCollection scollection = new AutoCompleteStringCollection();

            while (dr.Read())
            {
                string x = dr["vergiNumarasi"].ToString();
                string y = dr["sirketUnvan"].ToString();
                scollection.Add(x + " " + y);
            }


            vnTextBox.AutoCompleteCustomSource = scollection;

            con.Close();
        }

        List<TextBox> aliciBireyselComponentList = new List<TextBox>();
        List<TextBox> aliciSirketComponentList = new List<TextBox>();

        private void getirButton_Click(object sender, EventArgs e)
        {

            AliciSirket aliciSirket = new AliciSirket();
            aliciSirket = aliciSirketBilgileriGetir();

            if (aliciSirket == null)
            {
                MessageBox.Show("Şirket bilgileri getirilemedi. Textbox'da değer olduğundan emin misiniz?");
                return;
            }
            else
            {
                foreach (var item in aliciBireyselComponentList)
                {
                    item.Text = "";
                }

                foreach (var item in aliciSirketComponentList)
                {
                    switch (item.Name)
                    {
                        case "unvanTextBox":
                            item.Text = aliciSirket.SirketUnvan;
                            break;
                        case "vergiDairesiTextBox":
                            item.Text = aliciSirket.VergiDairesi;
                            break;
                        case "vnTextBox":
                            item.Text = aliciSirket.VergiNumarasi;
                            break;
                        case "telTextBox":
                            item.Text = aliciSirket.SirketTelefon;
                            break;
                        case "adresTextBox":
                            item.Text = aliciSirket.SirketAdres;
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        public bool BireyselAliciVarMi(string bireyselTc)
        {
            // Sorgu sonucu
            bool aliciVar = false;

            // Veritabanı bağlantısı ve komut nesnesi oluşturma
            using (SqlConnection con = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("SELECT COUNT(aliciTc) FROM alici_bireysel WHERE aliciTc = @aliciTc", con))
            {
                // Parametre ekleme
                cmd.Parameters.AddWithValue("@aliciTc", bireyselTc);

                // Bağlantıyı açma
                con.Open();

                // Sorguyu çalıştırma ve sonuç alma
                int count = (int)cmd.ExecuteScalar();

                // Sonucu değerlendirme
                aliciVar = (count > 0);
            }

            // Sonucu döndürme
            return aliciVar;
        }

        public bool SirketAliciVarMi(string sirketVergiNumarasi)
        {

            // Sorgu sonucu
            bool aliciVar = false;

            // Veritabanı bağlantısı ve komut nesnesi oluşturma
            using (SqlConnection con = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("SELECT COUNT(vergiNumarasi) FROM alici_sirket WHERE vergiNumarasi = @vergiNo", con))
            {
                // Parametre ekleme
                cmd.Parameters.AddWithValue("@vergiNo", sirketVergiNumarasi);

                // Bağlantıyı açma
                con.Open();

                // Sorguyu çalıştırma ve sonuç alma
                int count = (int)cmd.ExecuteScalar();

                // Sonucu değerlendirme
                aliciVar = (count > 0);
            }

            // Sonucu döndürme
            return aliciVar;
        }

        public AliciSirket aliciSirketBilgileriGetirFromComponents()
        {
            AliciSirket aliciSirket = new AliciSirket();

            foreach (var item in aliciSirketComponentList)
            {
                switch (item.Name)
                {
                    case "unvanTextBox":
                        aliciSirket.SirketUnvan = item.Text;
                        break;
                    case "vergiDairesiTextBox":
                        aliciSirket.VergiDairesi = item.Text;
                        break;
                    case "vnTextBox":
                        aliciSirket.VergiNumarasi = item.Text;
                        break;
                    case "telTextBox":
                        aliciSirket.SirketTelefon = item.Text;
                        break;
                    case "adresTextBox":
                        aliciSirket.SirketAdres = item.Text;
                        break;
                    default:
                        break;
                }
            }
            return aliciSirket;
        }


        public AliciSirket aliciSirketBilgileriGetir()
        {
            SqlConnection con = new SqlConnection(connectionString);

            con.Open();

            string vnTxt = vnTextBox.Text;

            SqlCommand cmd = null;
            if (vnTxt.Length >= 10)
            {
                cmd = new SqlCommand("select * from alici_sirket where vergiNumarasi like '" + vnTxt.Substring(0, 10) + "'", con);
            }
            else
            {
                //MessageBox.Show("Şirket bilgileri getirilemedi. Textbox'da değer olduğundan emin misiniz?");
                return null;
            }

            SqlDataReader dr = cmd.ExecuteReader();

            AliciSirket aliciSirket = new AliciSirket();

            while (dr.Read())
            {
                string vn = dr["vergiNumarasi"].ToString();
                aliciSirket.VergiNumarasi = vn;

                string sirketUnvan = dr["sirketUnvan"].ToString();
                aliciSirket.SirketUnvan = sirketUnvan;

                string vd = dr["vergiDairesi"].ToString();
                aliciSirket.VergiDairesi = vd;

                string sirketAdres = dr["sirketAdres"].ToString();
                aliciSirket.SirketAdres = sirketAdres;

                string sirketTel = dr["sirketTelefon"].ToString();
                aliciSirket.SirketTelefon = sirketTel;
            }

            con.Close();
            return aliciSirket;
        }

        public AliciBireysel aliciBilgileriGetirFromComponents()
        {
            AliciBireysel aliciBireysel = new AliciBireysel();

            foreach (var item in aliciBireyselComponentList)
            {
                switch (item.Name)
                {
                    case "adSoyadTextBox":
                        aliciBireysel.AdSoyad = item.Text;
                        break;
                    case "tcknTextBox":
                        aliciBireysel.AliciTc = item.Text;
                        break;
                    case "telTextBox":
                        aliciBireysel.AliciTelefon = item.Text;
                        break;
                    case "adresTextBox":
                        aliciBireysel.AliciAdres = item.Text;
                        break;
                    default:
                        break;
                }
            }
            return aliciBireysel;
        }

        public AliciBireysel aliciBilgileriGetir()
        {
            SqlConnection con = new SqlConnection(connectionString);

            con.Open();


            string tcknTxt = tcknTextBox.Text;
            SqlCommand cmd = null;
            if (tcknTxt.Length >= 11)
            {
                cmd = new SqlCommand("select * from alici_bireysel where aliciTc like '" + tcknTextBox.Text.Substring(0, 11) + "'", con);
            }
            else
            {
                return null;
            }

            SqlDataReader dr = cmd.ExecuteReader();

            AliciBireysel aliciBireysel = new AliciBireysel();

            while (dr.Read())
            {
                string aliciTc = dr["aliciTc"].ToString();
                aliciBireysel.AliciTc = aliciTc;

                string adSoyad = dr["adSoyad"].ToString();
                aliciBireysel.AdSoyad = adSoyad;

                string aliciAdres = dr["aliciAdres"].ToString();
                aliciBireysel.AliciAdres = aliciAdres;

                string aliciTel = dr["aliciTelefon"].ToString();
                aliciBireysel.AliciTelefon = aliciTel;
            }

            con.Close();

            return aliciBireysel;
        }

        private void getirBireyselButton_Click(object sender, EventArgs e)
        {
            foreach (var item in aliciSirketComponentList)
            {
                item.Text = "";
            }

            AliciBireysel aliciBireysel = new AliciBireysel();
            aliciBireysel = aliciBilgileriGetir();

            if (aliciBireysel == null)
            {
                MessageBox.Show("Bireysel müşteri bilgileri getirilemedi. Textbox'da değer olduğundan emin misiniz?");
                return;
            }
            else
            {
                foreach (var item in aliciBireyselComponentList)
                {
                    switch (item.Name)
                    {
                        case "adSoyadTextBox":
                            item.Text = aliciBireysel.AdSoyad;
                            break;
                        case "tcknTextBox":
                            item.Text = aliciBireysel.AliciTc;
                            break;
                        case "telTextBox":
                            item.Text = aliciBireysel.AliciTelefon;
                            break;
                        case "adresTextBox":
                            item.Text = aliciBireysel.AliciAdres;
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        private void killAllWordProcessButton_Click(object sender, EventArgs e)
        {
            KillAllWordProcesses();
        }

        public static void KillAllWordProcesses()
        {
            foreach (var process in Process.GetProcessesByName("WINWORD"))
            {
                try
                {
                    process.Kill();
                    process.WaitForExit(); // İsteğe bağlı: sürecin sonlanmasını bekler
                }
                catch
                {
                    // Burada hataları loglayabilir ya da görmezden gelebilirsiniz.
                }
            }
        }

        private void temizleButton_Click(object sender, EventArgs e)
        {
            string kaynakKlasor = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Word", "ihtimaller", "yedek"); // Kaynak klasörün yolu
            string hedefKlasor = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Word", "ihtimaller");
            KlasorIceriginiKopyala(kaynakKlasor, hedefKlasor);
        }

        static void KlasorIceriginiKopyala(string kaynakKlasor, string hedefKlasor)
        {
            try
            {
                // Hedef klasördeki tüm dosyaları sil
                if (Directory.Exists(hedefKlasor))
                {
                    foreach (string dosya in Directory.GetFiles(hedefKlasor))
                    {
                        File.Delete(dosya);
                    }
                }
                else
                {
                    // Hedef klasör yoksa oluştur
                    Directory.CreateDirectory(hedefKlasor);
                }

                // Kaynak klasördeki tüm dosyaları hedefe kopyala
                if (Directory.Exists(kaynakKlasor))
                {
                    foreach (string dosya in Directory.GetFiles(kaynakKlasor))
                    {
                        string dosyaAdi = Path.GetFileName(dosya);
                        string hedefDosya = Path.Combine(hedefKlasor, dosyaAdi);
                        File.Copy(dosya, hedefDosya, true); // true ile var olanı ez
                    }
                }
                else
                {
                    throw new DirectoryNotFoundException("Kaynak klasör bulunamadı.");
                }
            }
            catch (System.IO.IOException)
            {
                MessageBox.Show("Taslak olarak kullanılacak dosya hali hazırda açık. Lütfen açık olan belgelerinizi kapatıp, 'Kill All' yaptıktan sonra tekrar deneyiniz.");
                return;
            }
            //:

        }

        private void adSoyadTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void tcknTextBox_KeyUp(object sender, KeyEventArgs e)
        {

        }

        private void yeniSozlesmeButton_Click(object sender, EventArgs e)
        {
            yeniSozlesme();
        }
    }
}
