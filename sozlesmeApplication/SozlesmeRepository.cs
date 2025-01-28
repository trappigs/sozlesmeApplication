using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sozlesmeApplication
{
    internal class SozlesmeRepository
    {
        private readonly DatabaseHelper _databaseHelper;

        public SozlesmeRepository(DatabaseHelper databaseHelper)
        {
            _databaseHelper = databaseHelper;
        }
    }
}
