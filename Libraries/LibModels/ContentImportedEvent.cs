using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Libraries.LibModels
{
    public class ContentImportedEvent
    {
        public Guid JobId { get; set; }
        public string SourceChannel { get; set; } = string.Empty;
        public int ItemsImportedCount { get; set; }
        public int ItemsSkippedCount { get; set; }
        public DateTime CompletedAtUtc { get; set; } = DateTime.UtcNow;
    }
}
