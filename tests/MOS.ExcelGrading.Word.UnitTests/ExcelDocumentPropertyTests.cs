using System;
using System.Collections.Generic;
using System.Reflection;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Services;
using Xunit;

namespace MOS.ExcelGrading.Word.UnitTests
{
    public class ExcelDocumentPropertyTests
    {
        private const string CoreXmlWithContentStatus = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"">
    <cp:contentStatus>Draft</cp:contentStatus>
</cp:coreProperties>";

        private const string CoreXmlWithEmptyContentStatus = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"">
    <cp:contentStatus></cp:contentStatus>
</cp:coreProperties>";

        private const string CustomXmlWithStatus = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/custom-properties""
            xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"">
    <property fmtid=""{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"" pid=""2"" name=""Status"">
        <vt:lpwstr>Draft</vt:lpwstr>
    </property>
</Properties>";

        private static (bool IsPassed, string Message) EvaluateExcel(ExcelDocumentPropertyConfig config, Dictionary<string, string> xmlParts)
        {
            var service = typeof(XmlGradingRuleService);
            var packageType = service.GetNestedType("OfficePackage", BindingFlags.NonPublic)!;
            var package = Activator.CreateInstance(packageType, nonPublic: true)!;
            var parts = (Dictionary<string, string>)packageType.GetProperty("XmlParts")!.GetValue(package)!;

            foreach (var kvp in xmlParts)
            {
                parts[kvp.Key] = kvp.Value;
            }

            var method = service.GetMethod("EvaluateExcelDocumentProperty", BindingFlags.Static | BindingFlags.NonPublic)!;
            var outcome = method.Invoke(null, new object?[] { config, package })!;

            var isPassed = (bool)outcome.GetType().GetProperty("IsPassed")!.GetValue(outcome)!;
            var message = (string?)outcome.GetType().GetProperty("Message")!.GetValue(outcome) ?? string.Empty;
            return (isPassed, message);
        }

        [Fact]
        public void MatchesCoreXmlContentStatus_Passes()
        {
            var config = new ExcelDocumentPropertyConfig { PropertyName = "Status", ExpectedValue = "Draft", SourceFile = "" };
            var parts = new Dictionary<string, string> { ["docProps/core.xml"] = CoreXmlWithContentStatus };
            var (isPassed, message) = EvaluateExcel(config, parts);
            Assert.True(isPassed, message);
        }

        [Fact]
        public void MatchesCustomXmlStatus_Passes()
        {
            var config = new ExcelDocumentPropertyConfig { PropertyName = "Status", ExpectedValue = "Draft", SourceFile = "docProps/custom.xml" };
            var parts = new Dictionary<string, string> { ["docProps/custom.xml"] = CustomXmlWithStatus };
            var (isPassed, message) = EvaluateExcel(config, parts);
            Assert.True(isPassed, message);
        }

        [Fact]
        public void CoreXmlHasEmptyContentStatus_CustomXmlHasDraft_Passes()
        {
            var config = new ExcelDocumentPropertyConfig { PropertyName = "Status", ExpectedValue = "Draft", SourceFile = "" };
            var parts = new Dictionary<string, string> { ["docProps/core.xml"] = CoreXmlWithEmptyContentStatus, ["docProps/custom.xml"] = CustomXmlWithStatus };
            var (isPassed, message) = EvaluateExcel(config, parts);
            Assert.True(isPassed, message);
        }

        [Fact]
        public void SourceFileSetToCustomXml_CoreXmlHasDraft_StillPasses()
        {
            var config = new ExcelDocumentPropertyConfig { PropertyName = "Status", ExpectedValue = "Draft", SourceFile = "docProps/custom.xml" };
            var parts = new Dictionary<string, string> { ["docProps/core.xml"] = CoreXmlWithContentStatus };
            var (isPassed, message) = EvaluateExcel(config, parts);
            Assert.True(isPassed, message);
        }

        [Fact]
        public void CaseInsensitiveAndWhitespaceNormalized_Passes()
        {
            var config = new ExcelDocumentPropertyConfig { PropertyName = "Status", ExpectedValue = "Draft" };
            const string coreXmlWithWhitespace = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties""><cp:contentStatus>   draft   </cp:contentStatus></cp:coreProperties>";
            var parts = new Dictionary<string, string> { ["docProps/core.xml"] = coreXmlWithWhitespace };
            var (isPassed, message) = EvaluateExcel(config, parts);
            Assert.True(isPassed, message);
        }

        [Fact]
        public void WrongValue_FailsWithHelpfulMessage()
        {
            var config = new ExcelDocumentPropertyConfig { PropertyName = "Status", ExpectedValue = "Draft" };
            const string coreXmlWrong = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties""><cp:contentStatus>Final</cp:contentStatus></cp:coreProperties>";
            var parts = new Dictionary<string, string> { ["docProps/core.xml"] = coreXmlWrong };
            var (isPassed, message) = EvaluateExcel(config, parts);
            Assert.False(isPassed);
            Assert.Contains("Final", message);
            Assert.Contains("Draft", message);
        }

        [Fact]
        public void PropertyMissing_FailsWithMissingMessage()
        {
            var config = new ExcelDocumentPropertyConfig { PropertyName = "Status", ExpectedValue = "Draft" };
            const string coreXmlNoProp = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties""></cp:coreProperties>";
            var parts = new Dictionary<string, string> { ["docProps/core.xml"] = coreXmlNoProp };
            var (isPassed, message) = EvaluateExcel(config, parts);
            Assert.False(isPassed);
            Assert.Contains("Chưa thêm thuộc tính", message);
        }
    }
}
