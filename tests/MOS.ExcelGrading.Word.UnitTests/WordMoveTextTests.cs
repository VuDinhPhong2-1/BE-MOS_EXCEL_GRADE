using System;
using System.Collections.Generic;
using System.Reflection;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Services;
using Xunit;

namespace MOS.ExcelGrading.Word.UnitTests;

public class WordMoveTextTests
{
    private static string P(string text) => $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>";

    [Theory]
    [InlineData("Before|Moved|After", true)]
    [InlineData("Moved|Before|After", false)]
    [InlineData("Before|Moved|After|Moved", false)]
    [InlineData("Before|After", false)]
    [InlineData("Before||Moved|After", false)]
    [InlineData("Before|Moved|After|Before", false)]
    public void ChecksPositionAndDuplicates(string paragraphs, bool expected)
    {
        var body = string.Concat(Array.ConvertAll(paragraphs.Split('|'), P));
        Assert.Equal(expected, Evaluate(body, new WordMoveTextConfig { ExpectedText = "Moved", AfterText = "Before", BeforeText = "After" }));
    }

    [Theory]
    [InlineData("ignore", null, true)]
    [InlineData("default", null, false)]
    [InlineData("default", "Quote", true)]
    [InlineData("custom", "Quote", true)]
    [InlineData("custom", "Normal", false)]
    [InlineData("unknown", "Quote", false)]
    public void ChecksConfiguredOutputStyle(string mode, string? style, bool expected)
    {
        var body = P("Before") + "<w:p><w:pPr><w:pStyle w:val='Quote'/></w:pPr><w:r><w:t>Moved</w:t></w:r></w:p>";
        Assert.Equal(expected, Evaluate(body, new WordMoveTextConfig { ExpectedText = "Moved", AfterText = "Before", PasteMode = mode, ExpectedParagraphStyle = style }));
    }

    [Fact]
    public void RejectsOriginalLocation() => Assert.False(Evaluate(P("Before") + P("Moved"),
        new WordMoveTextConfig { ExpectedText = "Moved", AfterText = "Before", OriginalAfterText = "Before" }));

    [Fact]
    public void RejectsTableBetweenAnchorAndTarget() => Assert.False(Evaluate(P("Before") + "<w:tbl/>" + P("Moved"),
        new WordMoveTextConfig { ExpectedText = "Moved", AfterText = "Before" }));

    private static bool Evaluate(string body, WordMoveTextConfig config)
    {
        var service = typeof(XmlGradingRuleService);
        var packageType = service.GetNestedType("OfficePackage", BindingFlags.NonPublic)!;
        var package = Activator.CreateInstance(packageType, nonPublic: true)!;
        var parts = (Dictionary<string, string>)packageType.GetProperty("XmlParts")!.GetValue(package)!;
        parts["word/document.xml"] = "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body>" + body + "</w:body></w:document>";
        var outcome = service.GetMethod("EvaluateWordMoveText", BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, new object[] { config, package })!;
        return (bool)outcome.GetType().GetProperty("IsPassed")!.GetValue(outcome)!;
    }
}