using System.Text.Json.Serialization;

namespace Libraries.LibModels
{
    public class ContentItem
    {
        [JsonExtensionData]
        public Dictionary<string, object> Fields { get; set; } = new(StringComparer.OrdinalIgnoreCase);

        [JsonIgnore]
        public string? ExternalId => Fields.TryGetValue("ExternalId", out var value) ? value?.ToString() : null;

        [JsonIgnore]
        public int FieldCount => Fields.Count;
    }
}
