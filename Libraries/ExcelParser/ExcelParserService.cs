using ExcelDataReader;
using Libraries.LibModels;
using System.Text;

namespace Libraries.ExcelParser
{
    public class ExcelParserService : IExcelParserService
    {
        private enum SpreadsheetFormat { Csv, Xlsx, Xls }

        static ExcelParserService()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public async Task<ExcelParseResult> ParseStreamAsync(Stream stream, string fileName, CancellationToken cancellationToken)
        {
            // Buffer into memory so we can both sniff the format (magic bytes) and read the body.
            using var buffer = new MemoryStream();
            await stream.CopyToAsync(buffer, cancellationToken);
            var bytes = buffer.ToArray();

            var format = DetectFormat(fileName, bytes);

            List<string[]> rawRows = format switch
            {
                SpreadsheetFormat.Csv => ParseCsv(bytes),
                _ => ParseExcel(bytes)
            };

            return BuildResult(rawRows, format, cancellationToken);
        }

        private static SpreadsheetFormat DetectFormat(string fileName, byte[] bytes)
        {
            var extension = Path.GetExtension(fileName)?.ToLowerInvariant();
            if (extension == ".csv") return SpreadsheetFormat.Csv;
            if (extension == ".xlsx" || extension == ".xlsm") return SpreadsheetFormat.Xlsx;
            if (extension == ".xls") return SpreadsheetFormat.Xls;

            if (bytes.Length >= 4 && bytes[0] == 0x50 && bytes[1] == 0x4B)
                return SpreadsheetFormat.Xlsx;
            if (bytes.Length >= 8 && bytes[0] == 0xD0 && bytes[1] == 0xCF &&
                bytes[2] == 0x11 && bytes[3] == 0xE0)
                return SpreadsheetFormat.Xls;

            return SpreadsheetFormat.Csv;
        }

        private static List<string[]> ParseCsv(byte[] bytes)
        {
            var rows = new List<string[]>();
            using var reader = new StreamReader(new MemoryStream(bytes), Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            string? line;
            while ((line = reader.ReadLine()) != null)
                rows.Add(SplitCsvLine(line));
            return rows;
        }

        private static string[] SplitCsvLine(string line)
        {
            var values = new List<string>();
            var current = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"') { current.Append('"'); i++; }
                        else inQuotes = false;
                    }
                    else current.Append(c);
                }
                else if (c == '"') inQuotes = true;
                else if (c == ',') { values.Add(current.ToString()); current.Clear(); }
                else current.Append(c);
            }

            values.Add(current.ToString());
            return values.ToArray();
        }

        private static List<string[]> ParseExcel(byte[] bytes)
        {
            var rows = new List<string[]>();
            using var stream = new MemoryStream(bytes);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            do
            {
                while (reader.Read())
                {
                    var values = new string[reader.FieldCount];
                    for (int i = 0; i < reader.FieldCount; i++)
                        values[i] = reader.GetValue(i)?.ToString()?.Trim() ?? string.Empty;
                    rows.Add(values);
                }
            } while (reader.NextResult());

            return rows;
        }

        private static ExcelParseResult BuildResult(List<string[]> rawRows, SpreadsheetFormat format, CancellationToken cancellationToken)
        {
            var items = new List<ContentItem>();
            var errors = new List<string>();
            var seenExternalIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int duplicateRows = 0;
            int totalRows = 0;

            if (rawRows.Count == 0)
            {
                return new ExcelParseResult
                {
                    Items = items,
                    TotalRowsProcessed = 0,
                    DetectedFormat = format.ToString()
                };
            }

            // Row 1 is the header row; every column becomes a dynamic field name.
            var headers = BuildHeaders(rawRows[0]);

            for (int rowIndex = 1; rowIndex < rawRows.Count; rowIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var cols = rawRows[rowIndex];
                if (IsEmptyRow(cols)) continue;

                totalRows++;

                var fields = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                for (int c = 0; c < headers.Count; c++)
                {
                    var value = c < cols.Length ? cols[c]?.Trim() ?? string.Empty : string.Empty;
                    fields[headers[c]] = value;
                }

                var item = new ContentItem { Fields = fields };

                // Idempotency within the batch: drop repeated ExternalId values (when that column exists).
                var externalId = item.ExternalId;
                if (!string.IsNullOrWhiteSpace(externalId) && !seenExternalIds.Add(externalId))
                {
                    duplicateRows++;
                    continue;
                }

                items.Add(item);
            }

            return new ExcelParseResult
            {
                Items = items,
                TotalRowsProcessed = totalRows,
                DuplicateRows = duplicateRows,
                DetectedFormat = format.ToString(),
                RowErrors = errors
            };
        }

        private static List<string> BuildHeaders(string[] headerRow)
        {
            var headers = new List<string>();
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < headerRow.Length; i++)
            {
                var name = headerRow[i]?.Trim();
                if (string.IsNullOrWhiteSpace(name))
                    name = $"Column{i + 1}";

                // Guarantee unique field names even if the source repeats a header.
                var unique = name;
                int suffix = 2;
                while (!used.Add(unique))
                    unique = $"{name}_{suffix++}";

                headers.Add(unique);
            }

            return headers;
        }

        private static bool IsEmptyRow(string[] cols) =>
            cols == null || cols.All(c => string.IsNullOrWhiteSpace(c));
    }
}
