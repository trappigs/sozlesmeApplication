using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sozlesmeApplication
{
    internal class Satici
    {
        public string SaticiVn { get; set; } // NOT NULL
        public string SaticiUnvan { get; set; } // NOT NULL
        public string SaticiVergiDairesi { get; set; } // NOT NULL
        public string? SaticiMersisNumarasi { get; set; } // NULLABLE
        public string SaticiAdres { get; set; } // NOT NULL
    }
}
