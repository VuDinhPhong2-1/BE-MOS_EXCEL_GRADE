using System;
using System.Collections.Generic;
using System.Reflection;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Services;
using Xunit;

namespace MOS.ExcelGrading.Word.UnitTests
{
    public class SectionBreakBeforeTextTests
    {
        private const string Target = "<w:p><w:r><w:t>Seek Truth and Report It</w:t></w:r></w:p>";
        private const string Boundary = "<w:p><w:pPr><w:sectPr><w:cols/></w:sectPr></w:pPr></w:p>";

        [Theory]
        [InlineData("continuous", 2, true)]
        [InlineData("nextPage", 2, false)]
        [InlineData("continuous", 1, false)]
        public void UsesFollowingSectionProperties(string type, int columns, bool expected)
        {
            Assert.Equal(expected, Evaluate(Boundary + Target + Properties(type, columns)));
        }

        [Fact]
        public void UsesFirstFollowingSection_NotFinalBodySection()
        {
            Assert.True(Evaluate(Boundary + Target + "<w:p><w:pPr>" +
                Properties("continuous", 2) + "</w:pPr></w:p>" + Properties("nextPage", 1)));
        }

        [Fact]
        public void DoesNotAcceptCorrectPropertiesFromPreviousSection()
        {
            Assert.False(Evaluate("<w:p><w:pPr>" + Properties("continuous", 2) +
                "</w:pPr></w:p>" + Target + Properties("nextPage", 1)));
        }

        [Fact]
        public void MissingTypeDefaultsToNextPage()
        {
            Assert.False(Evaluate(Boundary + Target + "<w:sectPr><w:cols w:num='2'/></w:sectPr>"));
        }

        [Fact]
        public void MissingColumnsDefaultToOne()
        {
            Assert.True(Evaluate(Boundary + Target + "<w:sectPr><w:type w:val='continuous'/></w:sectPr>",
                new SectionBreakBeforeTextConfig { TargetText = "Seek Truth", ExpectedColumnCount = 1 }));
        }

        [Fact]
        public void RequiresBoundaryBeforeTarget()
        {
            Assert.False(Evaluate(Target + Properties("continuous", 2)));
        }

        [Fact]
        public void MissingFollowingSectionFails()
        {
            Assert.False(Evaluate(Boundary + Target));
        }

        [Fact]
        public void PrefersPreviousBoundaryWhenTargetAlsoEndsSection()
        {
            var targetEnd = "<w:p><w:pPr>" + Properties("continuous", 2) +
                "</w:pPr><w:r><w:t>Seek Truth and Report It</w:t></w:r></w:p>";
            Assert.True(Evaluate(Boundary + targetEnd + Properties("nextPage", 1)));
        }

        [Fact]
        public void NonImmediateModeUsesNearestBoundary()
        {
            var body = Boundary + "<w:p/>" + Target + Properties("continuous", 2);
            Assert.False(Evaluate(body));
            Assert.True(Evaluate(body, new SectionBreakBeforeTextConfig
            {
                TargetText = "Seek Truth", ExpectedColumnCount = 2, RequireImmediateBefore = false
            }));
        }

        [Fact]
        public void SameParagraphOptionRemainsExplicitlySupported()
        {
            var body = "<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>Seek Truth</w:t></w:r></w:p>" +
                Properties("continuous", 2);
            Assert.True(Evaluate(body));
            Assert.False(Evaluate(body, new SectionBreakBeforeTextConfig
            {
                TargetText = "Seek Truth", ExpectedColumnCount = 2, AllowSameParagraphSectPr = false
            }));
        }

        private static string Properties(string type, int columns) =>
            $"<w:sectPr><w:type w:val='{type}'/><w:cols w:num='{columns}'/></w:sectPr>";

        private static bool Evaluate(string body, SectionBreakBeforeTextConfig? config = null)
        {
            var serviceType = typeof(XmlGradingRuleService);
            var packageType = serviceType.GetNestedType("OfficePackage", BindingFlags.NonPublic)!;
            var package = Activator.CreateInstance(packageType, nonPublic: true)!;
            var parts = (Dictionary<string, string>)packageType.GetProperty("XmlParts")!.GetValue(package)!;
            parts["word/document.xml"] =
                "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body>" +
                body + "</w:body></w:document>";
            var method = serviceType.GetMethod("EvaluateSectionBreakBeforeText", BindingFlags.NonPublic | BindingFlags.Static)!;
            var outcome = method.Invoke(null, new[] { config ?? new SectionBreakBeforeTextConfig
            {
                TargetText = "Seek Truth", ExpectedColumnCount = 2
            }, package })!;
            return (bool)outcome.GetType().GetProperty("IsPassed")!.GetValue(outcome)!;
        }
    }
}