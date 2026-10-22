using Libraries.LibModels;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Libraries.ExcelParser
{
    public interface IExcelParserService
    {
        Task<ExcelParseResult> ParseStreamAsync(Stream stream, string fileName, CancellationToken cancellationToken);
    }
}
