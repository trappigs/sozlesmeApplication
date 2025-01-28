using Microsoft.Win32;
using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace sozlesmeApplication
{
    public class DatabaseConnectionHelper
    {
        public static string GetServerName()
        {
            try
            {
                string serverName = Environment.MachineName;
                RegistryView registryView = Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32;
                using (RegistryKey hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, registryView))
                {
                    RegistryKey instanceKey = hklm.OpenSubKey(@"SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL", false);
                    if (instanceKey != null && instanceKey.GetValueNames().Length > 0)
                    {
                        foreach (string instanceName in instanceKey.GetValueNames())
                        {
                            if (instanceName == "SQLEXPRESS")
                            {
                                return $"{serverName}\\{instanceName}";
                            }
                        }
                    }
                }
                return "Server Bulunamadı.";
            }
            catch (Exception)
            {
                return "Server Bulunamadı.";
            }
        }

        public static void SetupDatabase()
        {
            string serverName = GetServerName();
            if (serverName == "Server Bulunamadı.")
            {
                MessageBox.Show("SQL Server instance bulunamadı.");
                return;
            }

            string connectionString = $"Server={serverName};Integrated Security=True;";
            string databaseName = "sozlesme";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // Veritabanı oluştur
                    string createDatabaseQuery = $@"
IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = '{databaseName}')
BEGIN
    CREATE DATABASE [{databaseName}];
END
";
                    using (SqlCommand cmd = new SqlCommand(createDatabaseQuery, connection))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    // Veritabanına bağlan
                    connection.ChangeDatabase(databaseName);

                    // alici_sirket tablosu oluşturma
                    string createAliciSirketTableQuery = @"
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'alici_sirket')
BEGIN
    CREATE TABLE [dbo].[alici_sirket](
        [vergiNumarasi] VARCHAR(10) NOT NULL,
        [vergiDairesi] VARCHAR(100) NOT NULL,
        [sirketAdres] VARCHAR(500) NOT NULL,
        [sirketTelefon] VARCHAR(10) NOT NULL,
        [sirketUnvan] VARCHAR(4000) NOT NULL,
        CONSTRAINT [alici_sirket_verginumarasi_primary] PRIMARY KEY([vergiNumarasi])
    );
END
";
                    using (SqlCommand cmd = new SqlCommand(createAliciSirketTableQuery, connection))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    // alici_bireysel tablosu oluşturma
                    string createAliciBireyselTableQuery = @"
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'alici_bireysel')
BEGIN
    CREATE TABLE [dbo].[alici_bireysel](
        [aliciTc] VARCHAR(11) NOT NULL,
        [adSoyad] VARCHAR(1000) NOT NULL,
        [aliciAdres] VARCHAR(1000) NOT NULL,
        [aliciTelefon] VARCHAR(10) NOT NULL,
        CONSTRAINT [alici_bireysel_alicitc_primary] PRIMARY KEY([aliciTc])
    );
END
";
                    using (SqlCommand cmd = new SqlCommand(createAliciBireyselTableQuery, connection))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    // sozlesme tablosu oluşturma
                    string createSozlesmeTableQuery = @"
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'sozlesme')
BEGIN
    CREATE TABLE [dbo].[sozlesme](
        [sozlesmeId] BIGINT IDENTITY(1,1) NOT NULL,
        [sozlesmeyiOlusturanKullanici] VARCHAR(500) NOT NULL,
        [saticiUnvan] VARCHAR(500) NOT NULL,
        [aliciTc] VARCHAR(11) NULL,
        [aliciVN] VARCHAR(10) NULL,
        [tasinmazBedeli] DECIMAL(12, 2) NOT NULL,
        [pesinOdemeTutari] BIGINT NULL,
        [pesinOdemeTarihi] DATETIME NULL,
        [takasMalinCinsi] VARCHAR(4000) NULL,
        [takasMalinOzellikleri] VARCHAR(4000) NULL,
        [takasMalinTutari] BIGINT NULL,
        [takasMalinTeslimTarihi] DATETIME NULL,
        [taksitSayisi] TINYINT NULL,
        [taksitTutari] BIGINT NULL,
        [taksitBaslangicTarihi] DATETIME NULL,
        [taksitBitisTarihi] DATETIME NULL,
        [sozlesmeTarihi] DATETIME NOT NULL,
        CONSTRAINT [sozlesme_sozlesmeid_primary] PRIMARY KEY([sozlesmeId])
    );
END
";
                    using (SqlCommand cmd = new SqlCommand(createSozlesmeTableQuery, connection))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    // sozlesme tablosu için foreign key kısıtlamaları
                    // Önce tablo var mı yok mu kontrol et
                    string addForeignKeysQuery = @"
IF EXISTS (SELECT * FROM sys.tables WHERE name = 'sozlesme')
BEGIN
    IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE name = 'sozlesme_alicitc_foreign')
    BEGIN
        ALTER TABLE [dbo].[sozlesme] ADD CONSTRAINT [sozlesme_alicitc_foreign]
        FOREIGN KEY([aliciTc]) REFERENCES [dbo].[alici_bireysel]([aliciTc]);
    END

    IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE name = 'sozlesme_alicivn_foreign')
    BEGIN
        ALTER TABLE [dbo].[sozlesme] ADD CONSTRAINT [sozlesme_alicivn_foreign]
        FOREIGN KEY([aliciVN]) REFERENCES [dbo].[alici_sirket]([vergiNumarasi]);
    END
END
";
                    using (SqlCommand cmd = new SqlCommand(addForeignKeysQuery, connection))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veri tabanı ve/veya tablolar oluşturulurken hata: " + ex.Message);
            }
        }
    }
}
