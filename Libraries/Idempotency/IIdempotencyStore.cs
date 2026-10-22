namespace Libraries.Idempotency
{
    public interface IIdempotencyStore
    {
        Task<Guid?> TryGetJobIdAsync(string key, CancellationToken cancellationToken);
        Task SaveAsync(string key, Guid jobId, CancellationToken cancellationToken);
    }
}
