using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text.Json;
using MongoDB.Bson;
using MongoDB.Bson.Serialization;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Services;
using Xunit;

namespace MOS.ExcelGrading.Word.UnitTests
{
    public class SmartArtColorsTests
    {
        [Fact]
        public void ColorConfigurationSurvivesJsonAndBsonRoundTrip()
        {
            const string json = """
                {"taskId":"3","specialCondition":{"type":"wordSmartArtColors","score":28.4,
                "wordSmartArtColorsConfig":{"colorsFile":"word/diagrams/colors2.xml","expectedColorStyle":"accent5_6"}}}
                """;
            var task = JsonSerializer.Deserialize<TaskXmlRule>(json)!;
            typeof(XmlGradingRuleService)
                .GetMethod("NormalizeTaskForPersistence", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, new object[] { task });

            var document = task.ToBsonDocument();
            var config = document["specialCondition"]["wordSmartArtColorsConfig"].AsBsonDocument;
            Assert.Equal("accent5_6", config["expectedColorStyle"].AsString);
            Assert.Equal("word/diagrams/colors2.xml", config["colorsFile"].AsString);

            var restored = BsonSerializer.Deserialize<TaskXmlRule>(document);
            using var response = JsonDocument.Parse(JsonSerializer.Serialize(restored));
            var responseConfig = response.RootElement.GetProperty("specialCondition")
                .GetProperty("wordSmartArtColorsConfig");
            Assert.Equal("accent5_6", responseConfig.GetProperty("expectedColorStyle").GetString());
            Assert.Equal("word/diagrams/colors2.xml", responseConfig.GetProperty("colorsFile").GetString());
            Assert.Equal("wordSmartArtColors", restored.SpecialCondition!.Type);
            Assert.Equal(28.4m, restored.SpecialCondition.Score);
        }

        [Theory]
        [InlineData("accent5_6", true)]
        [InlineData("urn:microsoft.com/office/officeart/2005/8/colors/colorful5", true)]
        [InlineData("urn:microsoft.com/office/officeart/2005/8/colors/COLORFUL5#0", true)]
        [InlineData("urn:microsoft.com/office/officeart/2005/8/colors/colorful4", false)]
        [InlineData("colorful5_extra", false)]
        [InlineData("urn:microsoft.com/office/officeart/2005/8/colors/accent5_6", true)]
        [InlineData("urn:microsoft.com/office/officeart/2005/8/colors/accent5_6#0", true)]
        [InlineData("accent5_6_extra", false)]
        [InlineData("accent4_5", false)]
        [InlineData("", false)]
        public void MatchesStyleIdentifierWithoutRequiringDataOrDocument(string id, bool expected)
        {
            Assert.Equal(expected, Evaluate($"<dgm:colorsDef xmlns:dgm='http://schemas.openxmlformats.org/drawingml/2006/diagram' uniqueId='{id}'/>"));
        }

        [Fact]
        public void IncidentalTextDoesNotPass() => Assert.False(Evaluate(
            "<dgm:colorsDef xmlns:dgm='http://schemas.openxmlformats.org/drawingml/2006/diagram' uniqueId='wrong'><dgm:title val='accent5_6'/></dgm:colorsDef>"));

        [Theory]
        [InlineData(null)]
        [InlineData("not xml")]
        [InlineData("<colorsDef uniqueId='accent5_6'/>")]
        public void MissingOrInvalidPartFails(string? xml) => Assert.False(Evaluate(xml));

        [Theory]
        [InlineData("colorful5", true)]
        [InlineData("colorful4", false)]
        [InlineData("colorful5_extra", false)]
        public void CanonicalColorfulStyleMatchesExactly(string id, bool expected)
        {
            Assert.Equal(expected, Evaluate($"<dgm:colorsDef xmlns:dgm='http://schemas.openxmlformats.org/drawingml/2006/diagram' uniqueId='urn:microsoft.com/office/officeart/2005/8/colors/{id}'/>", "colorful5"));
        }

        private static bool Evaluate(string? xml, string expectedStyle = "accent5_6")
        {
            var service = typeof(XmlGradingRuleService);
            var packageType = service.GetNestedType("OfficePackage", BindingFlags.NonPublic)!;
            var package = Activator.CreateInstance(packageType, nonPublic: true)!;
            var parts = (Dictionary<string, string>)packageType.GetProperty("XmlParts")!.GetValue(package)!;
            if (xml != null) parts["word/diagrams/colors2.xml"] = xml;
            var method = service.GetMethod("EvaluateWordSmartArtColors", BindingFlags.NonPublic | BindingFlags.Static)!;
            var outcome = method.Invoke(null, new object[] {
                new WordSmartArtColorsConfig { ColorsFile = "word/diagrams/colors2.xml", ExpectedColorStyle = expectedStyle }, package })!;
            return (bool)outcome.GetType().GetProperty("IsPassed")!.GetValue(outcome)!;
        }
    }
}