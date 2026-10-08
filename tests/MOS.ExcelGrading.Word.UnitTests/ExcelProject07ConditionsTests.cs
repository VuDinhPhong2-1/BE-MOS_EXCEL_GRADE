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
        private static object CreateOfficePackageFromZip(string zipPath)
        {
            var service = typeof(XmlGradingRuleService);
            var packageType = service.GetNestedType("OfficePackage", BindingFlags.NonPublic)!;
            var package = Activator.CreateInstance(packageType, nonPublic: true)!;
            var partsProp = (Dictionary<string, string>)packageType.GetProperty("XmlParts")!.GetValue(package)!;
            var partNamesProp = (HashSet<string>)packageType.GetProperty("PartNames")!.GetValue(package)!;

            using var zip = System.IO.Compression.ZipFile.OpenRead(zipPath);
            foreach (var entry in zip.Entries)
            {
                var name = entry.FullName.Replace('\\', '/');
                partNamesProp.Add(name);
                if (name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                {
                    using var reader = new System.IO.StreamReader(entry.Open());
                    partsProp[name] = reader.ReadToEnd();
                }
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

        [Fact]
        public void EvaluateExcelProject07_AllTasks_RealExamAndAnswerFiles()
        {
            var examPath = @"c:\Users\vup60\OneDrive\Máy tính\MOS E19 - Nen tang\EXAM 01\Project 07\2019_Excel_107_Bonuses.xlsx";
            var answerPath = @"c:\Users\vup60\OneDrive\Máy tính\Project07 - Đáp án.xlsx";

            if (!System.IO.File.Exists(examPath) || !System.IO.File.Exists(answerPath))
            {
                return;
            }

            var examPackage = CreateOfficePackageFromZip(examPath);
            var answerPackage = CreateOfficePackageFromZip(answerPath);

            // Task 1: F4 mixed reference with $15 locking row 15, autofilled F4:F11
            var task1Config = new ExcelFormulaReferencesConfig
            {
                WorksheetName = "Employee Bonuses",
                Cell = "F4",
                RequiredFormulaFragments = new List<string> { "$15" },
                ExpectedSharedRange = "F4:F11",
                RequireOnlyDefinedNameReferences = false
            };
            var (examT1Passed, _) = InvokeEvaluator("EvaluateExcelFormulaReferences", task1Config, examPackage);
            Assert.False(examT1Passed, "Exam file should fail Task 1 because F4 has relative reference without $15");

            var (ansT1Passed, ansT1Msg) = InvokeEvaluator("EvaluateExcelFormulaReferences", task1Config, answerPackage);
            Assert.True(ansT1Passed, $"Answer file must pass Task 1: {ansT1Msg}");

            // Task 2: G4 autofill down to G11
            var task2Config = new ExcelFormulaReferencesConfig
            {
                WorksheetName = "Employee Bonuses",
                Cell = "G11",
                ExpectedSharedRange = "G4:G11",
                RequireOnlyDefinedNameReferences = false
            };
            var (examT2Passed, _) = InvokeEvaluator("EvaluateExcelFormulaReferences", task2Config, examPackage);
            Assert.False(examT2Passed, "Exam file should fail Task 2 because G5:G11 are empty");

            var (ansT2Passed, ansT2Msg) = InvokeEvaluator("EvaluateExcelFormulaReferences", task2Config, answerPackage);
            Assert.True(ansT2Passed, $"Answer file must pass Task 2: {ansT2Msg}");

            // Task 3: Delete row containing salesperson Allen on worksheet Parts
            var task3Config = new ExcelTableRowDeleteConfig
            {
                WorksheetName = "Parts",
                DeletedText = "Allen",
                RequireAbsent = true,
                MatchWholeWord = true
            };
            var (examT3Passed, _) = InvokeEvaluator("EvaluateExcelTableRowDelete", task3Config, examPackage);
            Assert.False(examT3Passed, "Exam file should fail Task 3 because Allen still exists on sheet Parts");

            var (ansT3Passed, ansT3Msg) = InvokeEvaluator("EvaluateExcelTableRowDelete", task3Config, answerPackage);
            Assert.True(ansT3Passed, $"Answer file must pass Task 3: {ansT3Msg}");

            // Task 4: Table style White, Table Style Medium 1 on worksheet Parts
            var task4Config = new ExcelTableCreateConfig
            {
                WorksheetName = "Parts",
                ExpectedTableStyle = "TableStyleMedium1"
            };
            var (examT4Passed, _) = InvokeEvaluator("EvaluateExcelTableCreate", task4Config, examPackage);
            Assert.False(examT4Passed, "Exam file should fail Task 4 because table style is TableStyleLight15");

            var (ansT4Passed, ansT4Msg) = InvokeEvaluator("EvaluateExcelTableCreate", task4Config, answerPackage);
            Assert.True(ansT4Passed, $"Answer file must pass Task 4: {ansT4Msg}");
        }

    }
}