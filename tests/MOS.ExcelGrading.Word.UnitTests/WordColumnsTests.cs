using System;
using System.Collections.Generic;
using System.Reflection;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Services;
using Xunit;

namespace MOS.ExcelGrading.Word.UnitTests
{
    public class WordColumnsTests
    {
        private static string Section(int columns) => $"<w:sectPr><w:cols w:num='{columns}'/></w:sectPr>";
        private static string Paragraph(string text, int? columns = null) =>
            "<w:p>" + (columns.HasValue ? "<w:pPr>" + Section(columns.Value) + "</w:pPr>" : "") +
            $"<w:r><w:t>{text}</w:t></w:r></w:p>";

        [Fact]
        public void ExactRangePasses() => Assert.True(Evaluate(Paragraph("Before", 1) + Paragraph("Start") + Paragraph("End", 2) + Paragraph("After") + Section(1)));

        [Theory]
        [InlineData(1, false)]
        [InlineData(2, true)]
        [InlineData(3, false)]
        public void EverySectionMustMatch(int middleColumns, bool expected) =>
            Assert.Equal(expected, Evaluate(Paragraph("Start", 2) + Paragraph("Middle", middleColumns) + Paragraph("End") + Section(2)));

        [Fact]
        public void RejectsOverflowBefore() => Assert.False(Evaluate(Paragraph("Before") + Paragraph("Start") + Paragraph("End") + Section(2)));

        [Fact]
        public void RejectsOverflowAfter() => Assert.False(Evaluate(Paragraph("Start") + Paragraph("End") + Paragraph("After") + Section(2)));

        [Fact]
        public void RejectsTableOverflow() => Assert.False(Evaluate("<w:tbl/>" + Paragraph("Start") + Paragraph("End") + Section(2)));

        [Fact]
        public void MissingPropertiesFails() => Assert.False(Evaluate(Paragraph("Start") + Paragraph("End")));

        [Fact]
        public void ReversedRangeFails() => Assert.False(Evaluate(Paragraph("End") + Paragraph("Start") + Section(2)));

        [Fact]
        public void MissingTargetFails() => Assert.False(Evaluate(Paragraph("Start") + Section(2)));

        [Fact]
        public void OccurrenceSelectsCorrectParagraph() => Assert.True(Evaluate(Paragraph("Start", 1) + Paragraph("Start") + Paragraph("End") + Section(2), 2));

        [Fact]
        public void UnequalColumnsUseExplicitChildren() => Assert.True(Evaluate(Paragraph("Start") + Paragraph("End") + "<w:sectPr><w:cols w:equalWidth='0'><w:col w:w='1000'/><w:col w:w='2000'/></w:cols></w:sectPr>"));

        private static bool Evaluate(string body, int occurrence = 1)
        {
            var service = typeof(XmlGradingRuleService);
            var packageType = service.GetNestedType("OfficePackage", BindingFlags.NonPublic)!;
            var package = Activator.CreateInstance(packageType, nonPublic: true)!;
            var parts = (Dictionary<string, string>)packageType.GetProperty("XmlParts")!.GetValue(package)!;
            parts["word/document.xml"] = "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body>" + body + "</w:body></w:document>";
            var outcome = service.GetMethod("EvaluateWordColumns", BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null,
                new object[] { new WordColumnsConfig { StartText = "Start", EndText = "End", StartOccurrence = occurrence }, package })!;
            return (bool)outcome.GetType().GetProperty("IsPassed")!.GetValue(outcome)!;
        }
    }
}