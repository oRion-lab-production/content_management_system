namespace Libraries.Storage
{
    public interface IStorageLocator
    {
        string Root { get; }
        string NoSqlPath { get; }
        string FilePath { get; }
    }
}
