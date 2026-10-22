namespace Libraries.LibModels
{
    public class ExcelParseResult
    {
        public List<ContentItem> Items { get; set; } = new();
        public int TotalRowsProcessed { get; set; }
        public int DuplicateRows { get; set; }
        public string DetectedFormat { get; set; } = string.Empty;
        public List<string> RowErrors { get; set; } = new();
    }
}
