using MongoDB.Bson.Serialization.Attributes;
using System.Text.Json.Serialization;

namespace MOS.ExcelGrading.Core.Models;

[BsonIgnoreExtraElements]
public class WordMoveTextConfig
{
    [BsonElement("expectedText"), JsonPropertyName("expectedText")]
    public string? ExpectedText { get; set; }
    [BsonElement("afterText"), JsonPropertyName("afterText")]
    public string? AfterText { get; set; }
    [BsonElement("beforeText"), JsonPropertyName("beforeText")]
    public string? BeforeText { get; set; }
    [BsonElement("originalAfterText"), JsonPropertyName("originalAfterText")]
    public string? OriginalAfterText { get; set; }
    [BsonElement("originalBeforeText"), JsonPropertyName("originalBeforeText")]
    public string? OriginalBeforeText { get; set; }
    [BsonElement("pasteMode"), JsonPropertyName("pasteMode")]
    public string PasteMode { get; set; } = "ignore";
    [BsonElement("expectedParagraphStyle"), JsonPropertyName("expectedParagraphStyle")]
    public string? ExpectedParagraphStyle { get; set; }
}