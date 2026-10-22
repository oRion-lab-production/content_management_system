namespace Libraries.Storage
{
    public interface IUploadFileStore
    {
        Task<string> SaveAsync(Guid jobId, string fileName, byte[] content, CancellationToken cancellationToken);
    }
}
