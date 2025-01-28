using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sozlesmeApplication
{
    public class Sozlesme
    {
        public long SozlesmeId { get; set; } // NOT NULL
        public string SozlesmeyiOlusturanKullanici { get; set; } // NOT NULL
        public string SaticiUnvan { get; set; } // NOT NULL
        public string? AliciTc { get; set; } // NULLABLE
        public string? AliciVN { get; set; } // NULLABLE
        public decimal TasinmazBedeli { get; set; } // NULLABLE
        public decimal? PesinOdemeTutari { get; set; } // NULLABLE
        public DateTime? PesinOdemeTarihi { get; set; } // NULLABLE
        public string? TakasMalinCinsi { get; set; } // NULLABLE
        public string? TakasMalinOzellikleri { get; set; } // NULLABLE
        public decimal? TakasMalinTutari { get; set; } // NULLABLE
        public DateTime? TakasMalinTeslimTarihi { get; set; } // NULLABLE
        public byte? TaksitSayisi { get; set; } // NULLABLE
        public decimal? TaksitTutari { get; set; } // NULLABLE
        public DateTime? TaksitBaslangicTarihi { get; set; } // NULLABLE
        public DateTime? TaksitBitisTarihi { get; set; } // NULLABLE
        public DateTime SozlesmeTarihi { get; set; } // NULLABLE
    }

}
