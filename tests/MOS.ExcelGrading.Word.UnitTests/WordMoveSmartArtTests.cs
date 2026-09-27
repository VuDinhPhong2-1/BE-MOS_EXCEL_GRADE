using System;
using System.Collections.Generic;
using System.Reflection;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Services;
using Xunit;

namespace MOS.ExcelGrading.Word.UnitTests;

public class WordMoveSmartArtTests
{
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string D = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    private static string P(string text) => $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>";
    private static string Art(string id = "rId8", string placement = "inline") =>
        $"<w:p><w:r><w:drawing><wp:{placement}><a:graphic><a:graphicData><dgm:relIds r:dm='{id}'/></a:graphicData></a:graphic></wp:{placement}></w:drawing></w:r></w:p>";
    private static WordMoveSmartArtConfig Config() => new() { NodeText = "Ethics", AfterText = "Heading", BeforeText = "Code of Ethics" };

    [Fact]
    public void ResolvesRelationshipRatherThanAssumingData1() => Assert.True(Evaluate(P("Heading") + Art() + P("Code of Ethics")));

    [Theory]
    [InlineData("unknown")]
    public void RejectsUnsupportedPlacement(string placement) => Assert.False(Evaluate(P("Heading") + Art(placement: placement) + P("Code of Ethics")));

    [Fact]
    public void AcceptsFloatingAtDestination() => Assert.True(Evaluate(P("Heading") + Art(placement: "anchor") + P("Code of Ethics")));

    [Fact]
    public void RejectsFloatingAtWrongDestination() => Assert.False(Evaluate(P("Other") + Art(placement: "anchor") + P("Code of Ethics")));

    [Fact]
    public void RejectsFloatingAtOldLocation()
    {
        var config = Config();
        config.OriginalBeforeText = "Code of Ethics";
        Assert.False(Evaluate(P("Heading") + Art(placement: "anchor") + P("Code of Ethics"), config));
    }

    [Theory]
    [InlineData("inline")]
    [InlineData("anchor")]
    public void RejectsNestedAndSharedParagraphs(string placement)
    {
        var art = Art(placement: placement);
        Assert.False(Evaluate(P("Heading") + "<w:tbl><w:tr><w:tc>" + art + "</w:tc></w:tr></w:tbl>" + P("Code of Ethics")));
        Assert.False(Evaluate(P("Heading") + "<w:p><w:r><w:txbxContent>" + art + "</w:txbxContent></w:r></w:p>" + P("Code of Ethics")));
        Assert.False(Evaluate(P("Heading") + art.Replace("</w:p>", "<w:r><w:t>Shared text</w:t></w:r></w:p>") + P("Code of Ethics")));
    }

    [Fact]
    public void FloatingPreservesAdjacencyAndUniqueness()
    {
        var body = P("Heading") + Art(placement: "anchor") + P("Code of Ethics");
        Assert.False(Evaluate(body + Art()));
        Assert.False(Evaluate(body + P("Heading")));
        Assert.False(Evaluate(P("Heading") + "<w:p/>" + Art(placement: "anchor") + P("Code of Ethics")));
        Assert.False(Evaluate(body.Replace("<wp:anchor>", "<wp:anchor><wp:inline>").Replace("</wp:anchor>", "</wp:inline></wp:anchor>")));
    }

    [Fact]
    public void RejectsDuplicateObjectsEvenWithSeparateDataParts() => Assert.False(Evaluate(P("Heading") + Art() + P("Code of Ethics") + Art("rId9")));

    [Fact]
    public void RejectsDuplicateReferencesToSamePart() => Assert.False(Evaluate(P("Heading") + Art() + P("Code of Ethics") + Art()));

    [Theory]
    [InlineData("<w:p/>")]
    [InlineData("<w:tbl/>")]
    public void DoesNotSkipBlocks(string block) => Assert.False(Evaluate(P("Heading") + block + Art() + P("Code of Ethics")));

    [Fact]
    public void RejectsOldLocation()
    {
        var config = Config();
        config.OriginalAfterText = "Heading";
        Assert.False(Evaluate(P("Heading") + Art() + P("Code of Ethics"), config));
    }

    [Fact]
    public void MatchesAutomaticallyNumberedAnchorByText() => Assert.True(Evaluate(P("Heading") + Art() +
        "<w:p><w:pPr><w:numPr><w:ilvl w:val='0'/><w:numId w:val='1'/></w:numPr></w:pPr><w:r><w:t>Code of Ethics</w:t></w:r></w:p>"));

    [Fact]
    public void RejectsAmbiguousAnchor() => Assert.False(Evaluate(P("Heading") + Art() + P("Code of Ethics") + P("Heading")));

    [Fact]
    public void RejectsMissingRelationship() => Assert.False(Evaluate(P("Heading") + Art("missing") + P("Code of Ethics")));

    [Fact]
    public void DoesNotMatchPresentationNode() => Assert.False(Evaluate(P("Heading") + Art() + P("Code of Ethics"), nodeType: "pres"));

    [Fact]
    public void RejectsWrongNodeText()
    {
        var config = Config();
        config.NodeText = "Other";
        Assert.False(Evaluate(P("Heading") + Art() + P("Code of Ethics"), config));
    }

    [Fact]
    public void RequiresDestination()
    {
        var config = Config();
        config.AfterText = config.BeforeText = null;
        Assert.False(Evaluate(P("Heading") + Art() + P("Code of Ethics"), config));
    }

    private static bool Evaluate(string body, WordMoveSmartArtConfig? config = null, string nodeType = "node")
    {
        var service = typeof(XmlGradingRuleService);
        var packageType = service.GetNestedType("OfficePackage", BindingFlags.NonPublic)!;
        var package = Activator.CreateInstance(packageType, nonPublic: true)!;
        var parts = (Dictionary<string, string>)packageType.GetProperty("XmlParts")!.GetValue(package)!;
        parts["word/document.xml"] = $"<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' xmlns:dgm='{D}' xmlns:r='{R}'><w:body>{body}</w:body></w:document>";
        parts["word/_rels/document.xml.rels"] = $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId8' Type='{R}/diagramData' Target='diagrams/data7.xml'/><Relationship Id='rId9' Type='{R}/diagramData' Target='diagrams/data8.xml'/></Relationships>";
        parts["word/diagrams/data7.xml"] = $"<dgm:dataModel xmlns:dgm='{D}' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'><dgm:ptLst><dgm:pt type='{nodeType}'><dgm:t><a:p><a:r><a:t>Ethics</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst></dgm:dataModel>";
        parts["word/diagrams/data8.xml"] = parts["word/diagrams/data7.xml"];
        var outcome = service.GetMethod("EvaluateWordMoveSmartArt", BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, new object[] { config ?? Config(), package })!;
        return (bool)outcome.GetType().GetProperty("IsPassed")!.GetValue(outcome)!;
    }
}