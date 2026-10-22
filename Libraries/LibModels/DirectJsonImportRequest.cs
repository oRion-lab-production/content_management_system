using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Libraries.LibModels
{
    public class DirectJsonImportRequest
    {
        public string SourceSystem { get; set; } = string.Empty;
        public List<ContentItem> Items { get; set; } = new();
    }
}
