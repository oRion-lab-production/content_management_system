using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Libraries.LibModels
{
    public class MigrationThresholds
    {
        public int MaxJsonItems { get; set; }
        public int MaxJsonSizeMb { get; set; }
        public int MaxExcelRows { get; set; }
        public int MaxExcelSizeMb { get; set; }
    }
}
