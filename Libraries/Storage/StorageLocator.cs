using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;

namespace Libraries.Storage
{
    public class StorageLocator : IStorageLocator
    {
        public string Root { get; }
        public string NoSqlPath { get; }
        public string FilePath { get; }

        public StorageLocator(IHostEnvironment environment, IConfiguration configuration)
        {
            // Default to <content root>/storage, overridable via configuration for deployment.
            var configuredRoot = configuration["StorageSettings:RootPath"];
            Root = string.IsNullOrWhiteSpace(configuredRoot)
                ? Path.Combine(environment.ContentRootPath, "storage")
                : configuredRoot;

            NoSqlPath = Path.Combine(Root, "nosql");
            FilePath = Path.Combine(Root, "file");

            Directory.CreateDirectory(NoSqlPath);
            Directory.CreateDirectory(FilePath);
        }
    }
}
