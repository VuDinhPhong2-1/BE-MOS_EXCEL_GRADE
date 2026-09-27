using System;
using System.Collections.Generic;
using System.Reflection;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Services;
using Xunit;

namespace MOS.ExcelGrading.Word.UnitTests
{
    public class SmartArtOrderTests
    {
        [Theory]
        [InlineData(3, true, true)]
        [InlineData(2, true, false)]
        [InlineData(0, true, false)]
        [InlineData(2, false, true)]
        public void LastNodeUsesConnectionOrderNotPointStorage(int targetOrder, bool requireLast, bool expected)
        {
            var points = "<dgm:pt modelId='root' type='doc'/>";
            var edges = "";
            for (var i = 3; i >= 0; i--)
            {
                var text = i == targetOrder ? "Be Accountable and Transparent" : "Other";
                points += $"<dgm:pt modelId='n{i}' type='node'><dgm:t><a:p><a:r><a:t>{text}</a:t></a:r></a:p></dgm:t></dgm:pt>";
                edges += $"<dgm:cxn type='parOf' srcId='root' destId='n{i}' srcOrd='{i}'/>";
            }
            Assert.Equal(expected, Evaluate(points, edges, requireLast));
        }

        [Fact]
        public void MissingConnectionsFailClosed()
        {
            Assert.False(Evaluate("<dgm:pt modelId='root' type='doc'/>", "", true));
        }

        [Theory]
        [InlineData(0, true)]
        [InlineData(1, true)]
        [InlineData(2, true)]
        [InlineData(3, true)]
        [InlineData(1, false)]
        public void FullOrderChecksEveryPosition(int targetIndex, bool correct)
        {
            var expected = new List<string> { "First", "Second", "Third", "Fourth" };
            expected[targetIndex] = "Be Accountable and Transparent";
            var actual = new List<string>(expected);
            if (!correct) (actual[1], actual[2]) = (actual[2], actual[1]);
            var points = "<dgm:pt modelId='root' type='doc'/>";
            var edges = "";
            for (var i = 3; i >= 0; i--)
            {
                points += $"<dgm:pt modelId='n{i}'><dgm:t><a:p><a:r><a:t>{actual[i]}</a:t></a:r></a:p></dgm:t></dgm:pt>";
                edges += $"<dgm:cxn type='parOf' srcId='root' destId='n{i}' srcOrd='{i}'/>";
            }
            // Full order supersedes the legacy last-node requirement.
            Assert.Equal(correct, Evaluate(points, edges, true, expected));
            Assert.False(Evaluate(points, edges, false, new List<string> { "First" }));
            Assert.False(Evaluate(points, "", false, expected));
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void ShapeCountIgnoresNonContentPointsWithText(bool omitNodeType, bool checkOrder)
        {
            var points = "<dgm:pt modelId='root' type='doc'><dgm:t/></dgm:pt>";
            var edges = "";
            var texts = new List<string> { "First", "Second", "Third", "Be Accountable and Transparent" };
            for (var i = 0; i < texts.Count; i++)
            {
                var type = omitNodeType ? "" : " type='node'";
                points += $"<dgm:pt modelId='n{i}'{type}><dgm:t><a:p><a:r><a:t>{texts[i]}</a:t></a:r></a:p></dgm:t></dgm:pt>";
                edges += $"<dgm:cxn type='parOf' srcId='root' destId='n{i}' srcOrd='{i}'/>";
            }
            foreach (var type in new[] { "pres", "parTrans", "sibTrans" })
                points += $"<dgm:pt modelId='{type}' type='{type}'><dgm:t><a:p><a:r><a:t>Display text</a:t></a:r></a:p></dgm:t></dgm:pt>";

            Assert.True(Evaluate(points, edges, checkOrder, checkOrder ? texts : null));
            Assert.False(Evaluate(points, edges, checkOrder, checkOrder ? texts : null, 3));
            Assert.False(Evaluate(points, edges, checkOrder, checkOrder ? texts : null, 5));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void ShapeCountIncludesContentNodesWithoutText(bool emptyTextElement)
        {
            var points = "<dgm:pt modelId='n0'><dgm:t><a:p><a:r><a:t>Be Accountable and Transparent</a:t></a:r></a:p></dgm:t></dgm:pt>";
            for (var i = 1; i < 4; i++)
                points += $"<dgm:pt modelId='n{i}'>{(emptyTextElement ? "<dgm:t/>" : "")}</dgm:pt>";

            Assert.True(Evaluate(points, "", false));
            Assert.False(Evaluate(points, "", false, expectedShapeCount: 3));
        }

        private static bool Evaluate(string points, string edges, bool requireLast, List<string>? expectedTexts = null, int expectedShapeCount = 4)
        {
            var service = typeof(XmlGradingRuleService);
            var packageType = service.GetNestedType("OfficePackage", BindingFlags.NonPublic)!;
            var package = Activator.CreateInstance(packageType, nonPublic: true)!;
            var parts = (Dictionary<string, string>)packageType.GetProperty("XmlParts")!.GetValue(package)!;
            parts["word/document.xml"] = "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body/></w:document>";
            parts["word/diagrams/data1.xml"] = $"<dgm:dataModel xmlns:dgm='http://schemas.openxmlformats.org/drawingml/2006/diagram' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'><dgm:ptLst>{points}</dgm:ptLst><dgm:cxnLst>{edges}</dgm:cxnLst></dgm:dataModel>";
            var config = new WordSmartArtConfig
            {
                ExpectedShapeCount = expectedShapeCount,
                ExpectedNodeTexts = expectedTexts,
                ExpectedText = "Be Accountable and Transparent",
                ExpectedLastNodeText = requireLast ? "Be Accountable and Transparent" : null
            };
            var outcome = service.GetMethod("EvaluateWordSmartArt", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, new object[] { config, package })!;
            return (bool)outcome.GetType().GetProperty("IsPassed")!.GetValue(outcome)!;
        }
    }
}