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

        private static bool Evaluate(string points, string edges, bool requireLast)
        {
            var service = typeof(XmlGradingRuleService);
            var packageType = service.GetNestedType("OfficePackage", BindingFlags.NonPublic)!;
            var package = Activator.CreateInstance(packageType, nonPublic: true)!;
            var parts = (Dictionary<string, string>)packageType.GetProperty("XmlParts")!.GetValue(package)!;
            parts["word/document.xml"] = "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body/></w:document>";
            parts["word/diagrams/data1.xml"] = $"<dgm:dataModel xmlns:dgm='http://schemas.openxmlformats.org/drawingml/2006/diagram' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'><dgm:ptLst>{points}</dgm:ptLst><dgm:cxnLst>{edges}</dgm:cxnLst></dgm:dataModel>";
            var config = new WordSmartArtConfig
            {
                ExpectedShapeCount = 4,
                ExpectedText = "Be Accountable and Transparent",
                ExpectedLastNodeText = requireLast ? "Be Accountable and Transparent" : null
            };
            var outcome = service.GetMethod("EvaluateWordSmartArt", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, new object[] { config, package })!;
            return (bool)outcome.GetType().GetProperty("IsPassed")!.GetValue(outcome)!;
        }
    }
}