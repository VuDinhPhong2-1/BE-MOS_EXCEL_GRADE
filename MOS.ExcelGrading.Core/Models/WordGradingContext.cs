using System.Xml.Linq;

namespace MOS.ExcelGrading.Core.Models
{
    public sealed class WordGradingContext
    {
        public int ProjectNumber { get; init; }
        public string SourceFileName { get; init; } = string.Empty;
        public byte[] PackageBytes { get; init; } = Array.Empty<byte>();
        public bool HasMainDocumentPart { get; init; }
        public int PartCount { get; init; }
        public HashSet<string> Entries { get; init; } = new(StringComparer.OrdinalIgnoreCase);
        public XDocument? MainDocumentXml { get; init; }
        public XDocument? CorePropertiesXml { get; init; }
        public XDocument? NumberingXml { get; init; }
        public XDocument? DocumentRelationshipsXml { get; init; }
        public Dictionary<string, WordDocumentRelationship> DocumentRelationships { get; init; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, XDocument> XmlParts { get; init; } = new(StringComparer.OrdinalIgnoreCase);

        public bool ContainsEntry(string entryName) => Entries.Contains(entryName);

        public bool TryGetDocumentRelationship(string relationshipId, out WordDocumentRelationship relationship)
        {
            return DocumentRelationships.TryGetValue(relationshipId, out relationship!);
        }

        public bool TryGetXmlPart(string entryName, out XDocument part)
        {
            var normalized = NormalizeEntryName(entryName);
            return XmlParts.TryGetValue(normalized, out part!);
        }

        private static string NormalizeEntryName(string entryName)
        {
            return (entryName ?? string.Empty).Replace("\\", "/", StringComparison.Ordinal);
        }
    }

    public sealed class WordDocumentRelationship
    {
        public string Id { get; init; } = string.Empty;
        public string Type { get; init; } = string.Empty;
        public string Target { get; init; } = string.Empty;
    }
}
