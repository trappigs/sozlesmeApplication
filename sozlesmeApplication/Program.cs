using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace sozlesmeApplication
{
    internal static class Program
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        [STAThread]
        static void Main()
        {
            // "Logs" klasörünün varlığını kontrol et ve yoksa oluştur
            string logsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
            if (!Directory.Exists(logsPath))
            {
                Directory.CreateDirectory(logsPath);
            }

            // Global hata yöneticilerini yapılandır
            Application.ThreadException += (sender, e) =>
            {
                Logger.Error(e.Exception, "UI iş parçacığında beklenmeyen bir hata oluştu.");
                MessageBox.Show("Hata: " + e.Exception.ToString() + "\nHatanın ayrıntıları: " + e.Exception.StackTrace);
                MessageBox.Show("Bir hata oluştu. Lütfen destek ekibiyle iletişime geçin.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };

            AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
            {
                var exception = e.ExceptionObject as Exception;
                Logger.Error(exception, "Genel uygulama alanında beklenmeyen bir hata oluştu.");
                MessageBox.Show("Hata: " + exception.ToString() + "\nHatanın ayrıntıları: " + exception.StackTrace);
                MessageBox.Show("Kritik bir hata oluştu. Uygulama kapatılacak.", "Kritik Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            };

            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new GirisForm());
        }
    }
}
