using Libraries.LibModels;

namespace Libraries.NoSqlDocumentRepositorys
{
    public interface INoSqlDocumentRepository
    {
        Task<DocumentSaveResult> SaveDocumentsAsync(Guid jobId, IEnumerable<ContentItem> items, string sourceChannel, CancellationToken cancellationToken);
    }
}
