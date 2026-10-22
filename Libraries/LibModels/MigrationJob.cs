using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Libraries.LibModels
{
    public class MigrationJob
    {
        public Guid JobId { get; set; }
        public string SourceChannel { get; set; } = string.Empty; // "JSON_API" or "EXCEL_UPLOAD"
        public List<ContentItem> Items { get; set; } = new();
        public string? NotificationEmail { get; set; }
        public DateTime CreatedAtUtc { get; set; } = DateTime.UtcNow;
    }
}
