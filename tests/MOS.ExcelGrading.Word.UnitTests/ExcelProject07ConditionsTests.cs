using System;
using System.Collections.Generic;
using System.Reflection;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Services;
using Xunit;

namespace MOS.ExcelGrading.Word.UnitTests
{
    public class ExcelProject07ConditionsTests
    {
        private static object CreateOfficePackage(Dictionary<string, string> xmlParts)
        {
            var service = typeof(XmlGradingRuleService);
            var packageType = service.GetNestedType("OfficePackage", BindingFlags.NonPublic)!;
            var package = Activator.CreateInstance(packageType, nonPublic: true)!;
            var partsProp = (Dictionary<string, string>)packageType.GetProperty("XmlParts")!.GetValue(package)!;
            foreach (var kvp in xmlParts)
            {
                partsProp[kvp.Key] = kvp.Value;
            }
            return package;
        }

        private static (bool IsPassed, string Message) InvokeEvaluator(string methodName, object? config, object package)
        {
            var service = typeof(XmlGradingRuleService);
            var method = service.GetMethod(methodName, BindingFlags.Static | BindingFlags.NonPublic)!;
            var outcome = method.Invoke(null, new object?[] { config, package })!;

            var isPassed = (bool)outcome.GetType().GetProperty("IsPassed")!.GetValue(outcome)!;
            var message = (string?)outcome.GetType().GetProperty("Message")!.GetValue(outcome) ?? string.Empty;
            return (isPassed, message);
        }

        [Fact]
        public void EvaluateExcelFormulaReferences_WithOnlyExpectedSharedRange_Passes()
        {
            var worksheetXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <sheetData>
        <row r=""4""><c r=""G4""><f t=""shared"" ref=""G4:G11"" si=""0"">F4*0.1</f><v>10</v></c></row>
        <row r=""11""><c r=""G11""><f t=""shared"" si=""0""/><v>20</v></c></row>
    </sheetData>
</worksheet>";

            var package = CreateOfficePackage(new Dictionary<string, string>
            {
                ["xl/worksheets/sheet1.xml"] = worksheetXml
            });

            var config = new ExcelFormulaReferencesConfig
            {
                SourceFile = "xl/worksheets/sheet1.xml",
                Cell = "G11",
                ExpectedSharedRange = "G4:G11",
                RequireOnlyDefinedNameReferences = false
            };

            var (isPassed, message) = InvokeEvaluator("EvaluateExcelFormulaReferences", config, package);
            Assert.True(isPassed, message);
        }

        [Fact]
        public void EvaluateExcelTableRowDelete_WhenDeletedTextMissing_Passes()
        {
            var worksheetXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <sheetData>
        <row r=""4""><c r=""A4"" t=""inlineStr""><is><t>Baker</t></is></c></row>
    </sheetData>
</worksheet>";

            var package = CreateOfficePackage(new Dictionary<string, string>
            {
                ["xl/worksheets/sheet1.xml"] = worksheetXml
            });

            var config = new ExcelTableRowDeleteConfig
            {
                SourceFile = "xl/worksheets/sheet1.xml",
                DeletedText = "Allen",
                MatchWholeWord = true
            };

            var (isPassed, message) = InvokeEvaluator("EvaluateExcelTableRowDelete", config, package);
            Assert.True(isPassed, message);
        }

        [Fact]
        public void EvaluateExcelTableRowDelete_WhenDeletedTextStillExists_Fails()
        {
            var worksheetXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <sheetData>
        <row r=""11""><c r=""A11"" t=""inlineStr""><is><t>Allen</t></is></c></row>
    </sheetData>
</worksheet>";

            var package = CreateOfficePackage(new Dictionary<string, string>
            {
                ["xl/worksheets/sheet1.xml"] = worksheetXml
            });

            var config = new ExcelTableRowDeleteConfig
            {
                SourceFile = "xl/worksheets/sheet1.xml",
                DeletedText = "Allen",
                MatchWholeWord = true
            };

            var (isPassed, _) = InvokeEvaluator("EvaluateExcelTableRowDelete", config, package);
            Assert.False(isPassed);
        }

        [Fact]
        public void EvaluateExcelSparkline_LineAtExpectedLocationAndDataRange_Passes()
        {
            var worksheetXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""
           xmlns:x14=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"">
    <extLst>
        <ext>
            <x14:sparklineGroups>
                <x14:sparklineGroup type=""line"">
                    <x14:sparklines>
                        <x14:sparkline>
                            <x14:f>Parts!B4:D4</x14:f>
                            <x14:sqref>F4</x14:sqref>
                        </x14:sparkline>
                    </x14:sparklines>
                </x14:sparklineGroup>
            </x14:sparklineGroups>
        </ext>
    </extLst>
</worksheet>";

            var package = CreateOfficePackage(new Dictionary<string, string>
            {
                ["xl/worksheets/sheet1.xml"] = worksheetXml
            });

            var config = new ExcelSparklineConfig
            {
                SourceFile = "xl/worksheets/sheet1.xml",
                SparklineType = "line",
                LocationRange = "F4",
                DataRange = "B4:D4"
            };

            var (isPassed, message) = InvokeEvaluator("EvaluateExcelSparkline", config, package);
            Assert.True(isPassed, message);
        }
    }
}