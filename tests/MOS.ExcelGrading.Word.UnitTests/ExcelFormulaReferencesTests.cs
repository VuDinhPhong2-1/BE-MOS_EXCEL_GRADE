using System;
using System.Collections.Generic;
using System.Reflection;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Services;
using Xunit;

namespace MOS.ExcelGrading.Word.UnitTests
{
    public class ExcelFormulaReferencesTests
    {
        private static string CreateWorksheetXml(string cellRef, string formula)
        {
            return $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <sheetData>
        <row r=""4"">
            <c r=""{cellRef}"">
                <f>{formula}</f>
                <v>Cross Country - Hardtail</v>
            </c>
        </row>
    </sheetData>
</worksheet>";
        }

        private static (bool IsPassed, string Message) EvaluateFormula(
            ExcelFormulaReferencesConfig config,
            string formulaXml)
        {
            var service = typeof(XmlGradingRuleService);
            var packageType = service.GetNestedType("OfficePackage", BindingFlags.NonPublic)!;
            var package = Activator.CreateInstance(packageType, nonPublic: true)!;
            var parts = (Dictionary<string, string>)packageType.GetProperty("XmlParts")!.GetValue(package)!;

            parts["xl/worksheets/sheet1.xml"] = CreateWorksheetXml(config.Cell ?? "B4", formulaXml);
            config.SourceFile = "xl/worksheets/sheet1.xml";

            var method = service.GetMethod("EvaluateExcelFormulaReferences", BindingFlags.Static | BindingFlags.NonPublic)!;
            var outcome = method.Invoke(null, new object?[] { config, package })!;

            var isPassed = (bool)outcome.GetType().GetProperty("IsPassed")!.GetValue(outcome)!;
            var message = (string?)outcome.GetType().GetProperty("Message")!.GetValue(outcome) ?? string.Empty;
            return (isPassed, message);
        }

        [Fact]
        public void FormulaWithAmpersandTwoSpaces_Passes()
        {
            var config = new ExcelFormulaReferencesConfig
            {
                Cell = "B4",
                RequiredReferences = new List<string> { "Description", "Style" },
                RequiredFormulaFragments = new List<string> { "\"  -  \"" },
                RequireOnlyDefinedNameReferences = true
            };

            var (isPassed, message) = EvaluateFormula(config, "[@Description] &amp; \"  -  \" &amp; [@Style]");
            Assert.True(isPassed, message);
        }

        [Fact]
        public void FormulaWithAmpersandTwoSpaces_WhenOneSpaceConfigured_PassesDueToFlexibleWhitespace()
        {
            var config = new ExcelFormulaReferencesConfig
            {
                Cell = "B4",
                RequiredReferences = new List<string> { "Description", "Style" },
                RequiredFormulaFragments = new List<string> { "\" - \"" },
                RequireOnlyDefinedNameReferences = true
            };

            var (isPassed, message) = EvaluateFormula(config, "[@Description] &amp; \"  -  \" &amp; [@Style]");
            Assert.True(isPassed, message);
        }

        [Fact]
        public void FormulaWithAmpersandOneSpace_WhenTwoSpacesConfigured_PassesDueToFlexibleWhitespace()
        {
            var config = new ExcelFormulaReferencesConfig
            {
                Cell = "B4",
                RequiredReferences = new List<string> { "Description", "Style" },
                RequiredFormulaFragments = new List<string> { "\"  -  \"" },
                RequireOnlyDefinedNameReferences = true
            };

            var (isPassed, message) = EvaluateFormula(config, "[@Description] &amp; \" - \" &amp; [@Style]");
            Assert.True(isPassed, message);
        }

        [Fact]
        public void FormulaWithConcatTwoSpaces_Passes()
        {
            var config = new ExcelFormulaReferencesConfig
            {
                Cell = "B4",
                RequiredReferences = new List<string> { "Description", "Style" },
                RequiredFormulaFragments = new List<string> { "\"  -  \"" },
                RequireOnlyDefinedNameReferences = true
            };

            var (isPassed, message) = EvaluateFormula(config, "CONCAT([@Description], \"  -  \", [@Style])");
            Assert.True(isPassed, message);
        }

        [Fact]
        public void FormulaWithXlfnConcatTwoSpaces_Passes()
        {
            var config = new ExcelFormulaReferencesConfig
            {
                Cell = "B4",
                RequiredReferences = new List<string> { "Description", "Style" },
                RequiredFormulaFragments = new List<string> { "\"  -  \"" },
                RequireOnlyDefinedNameReferences = true
            };

            var (isPassed, message) = EvaluateFormula(config, "_xlfn.CONCAT([@Description],\"  -  \",[@Style])");
            Assert.True(isPassed, message);
        }

        [Fact]
        public void FormulaWithConcatenateTwoSpaces_Passes()
        {
            var config = new ExcelFormulaReferencesConfig
            {
                Cell = "B4",
                RequiredReferences = new List<string> { "Description", "Style" },
                RequiredFormulaFragments = new List<string> { "\"  -  \"" },
                RequireOnlyDefinedNameReferences = true
            };

            var (isPassed, message) = EvaluateFormula(config, "CONCATENATE([@Description], \"  -  \", [@Style])");
            Assert.True(isPassed, message);
        }

        [Fact]
        public void FormulaWithDirectCellReferences_Fails()
        {
            var config = new ExcelFormulaReferencesConfig
            {
                Cell = "B4",
                RequiredReferences = new List<string> { "Description", "Style" },
                RequiredFormulaFragments = new List<string> { "\"  -  \"" },
                RequireOnlyDefinedNameReferences = true
            };

            var (isPassed, _) = EvaluateFormula(config, "A4 &amp; \"  -  \" &amp; C4");
            Assert.False(isPassed);
        }

        [Fact]
        public void FormulaWithNoSpacesAroundHyphen_Fails()
        {
            var config = new ExcelFormulaReferencesConfig
            {
                Cell = "B4",
                RequiredReferences = new List<string> { "Description", "Style" },
                RequiredFormulaFragments = new List<string> { "\"  -  \"" },
                RequireOnlyDefinedNameReferences = true
            };

            var (isPassed, _) = EvaluateFormula(config, "[@Description] &amp; \"-\" &amp; [@Style]");
            Assert.False(isPassed);
        }

        [Fact]
        public void FormulaMissingStyleReference_Fails()
        {
            var config = new ExcelFormulaReferencesConfig
            {
                Cell = "B4",
                RequiredReferences = new List<string> { "Description", "Style" },
                RequiredFormulaFragments = new List<string> { "\"  -  \"" },
                RequireOnlyDefinedNameReferences = true
            };

            var (isPassed, _) = EvaluateFormula(config, "[@Description] &amp; \"  -  \"");
            Assert.False(isPassed);
        }

        [Fact]
        public void FunctionCheckWithXlfnPrefix_Passes()
        {
            var config = new ExcelFormulaReferencesConfig
            {
                Cell = "B4",
                RequiredReferences = new List<string> { "Description", "Style" },
                RequiredFunctions = new List<string> { "CONCAT" },
                RequiredFormulaFragments = new List<string> { "\"  -  \"" },
                RequireOnlyDefinedNameReferences = true
            };

            var (isPassed, message) = EvaluateFormula(config, "_xlfn.CONCAT([@Description],\"  -  \",[@Style])");
            Assert.True(isPassed, message);
        }
    }
}