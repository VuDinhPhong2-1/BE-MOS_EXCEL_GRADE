using System;
using System.IO;
using System.Text.Json;
using MOS.ExcelGrading.Core.Models;
using Xunit;

namespace MOS.ExcelGrading.Api.UnitTests
{
    public class XmlGradingProject07RuleTests
    {
        [Fact]
        public void Project07Rules_Json_LoadsCorrectly_AndPassesAllValidations()
        {
            // Arrange: locate project07.json
            var possiblePaths = new[]
            {
                Path.Combine(AppContext.BaseDirectory, "project07.json"),
                Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "tools", "Project07RuleSeeder", "project07.json"),
                Path.Combine(Directory.GetCurrentDirectory(), "tools", "Project07RuleSeeder", "project07.json"),
                "/Users/stark/Documents/GitHub/MOS_Grader/BE-MOS_EXCEL_GRADE/tools/Project07RuleSeeder/project07.json"
            };

            string? jsonPath = null;
            foreach (var p in possiblePaths)
            {
                if (File.Exists(p))
                {
                    jsonPath = p;
                    break;
                }
            }

            Assert.NotNull(jsonPath);
            var json = File.ReadAllText(jsonPath);
            var project = JsonSerializer.Deserialize<GradingRuleProject>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            Assert.NotNull(project);
            Assert.Equal("project07", project.ProjectCode);
            Assert.Equal(142.0m, project.MaxScore);
            Assert.Equal(5, project.Tasks.Count);

            // Assert T1: Mixed reference formula in F4, autofill F4:F11
            var t1 = project.Tasks.Find(t => t.TaskId == "P07-T1");
            Assert.NotNull(t1);
            Assert.Equal(28.4m, t1.MaxScore);
            Assert.Equal(SpecialConditionTypes.ExcelFormulaReferences, t1.SpecialCondition?.Type);
            Assert.NotNull(t1.SpecialCondition?.ExcelFormulaReferencesConfig);
            Assert.Equal("Employee Bonuses", t1.SpecialCondition.ExcelFormulaReferencesConfig.WorksheetName);
            Assert.Equal("F4", t1.SpecialCondition.ExcelFormulaReferencesConfig.Cell);
            Assert.Equal("F4:F11", t1.SpecialCondition.ExcelFormulaReferencesConfig.ExpectedSharedRange);
            Assert.False(t1.SpecialCondition.ExcelFormulaReferencesConfig.RequireOnlyDefinedNameReferences);

            // Assert T2: Autofill formula in G4 down to G11
            var t2 = project.Tasks.Find(t => t.TaskId == "P07-T2");
            Assert.NotNull(t2);
            Assert.Equal(28.4m, t2.MaxScore);
            Assert.Equal(SpecialConditionTypes.ExcelFormulaReferences, t2.SpecialCondition?.Type);
            Assert.NotNull(t2.SpecialCondition?.ExcelFormulaReferencesConfig);
            Assert.Equal("G11", t2.SpecialCondition.ExcelFormulaReferencesConfig.Cell);
            Assert.Equal("G4:G11", t2.SpecialCondition.ExcelFormulaReferencesConfig.ExpectedSharedRange);

            // Assert T3: Delete row containing salesperson Allen on sheet Parts
            var t3 = project.Tasks.Find(t => t.TaskId == "P07-T3");
            Assert.NotNull(t3);
            Assert.Equal(28.4m, t3.MaxScore);
            Assert.Equal(SpecialConditionTypes.ExcelTableRowDelete, t3.SpecialCondition?.Type);
            Assert.NotNull(t3.SpecialCondition?.ExcelTableRowDeleteConfig);
            Assert.Equal("Parts", t3.SpecialCondition.ExcelTableRowDeleteConfig.WorksheetName);
            Assert.Equal("Allen", t3.SpecialCondition.ExcelTableRowDeleteConfig.DeletedText);
            Assert.True(t3.SpecialCondition.ExcelTableRowDeleteConfig.RequireAbsent);

            // Assert T4: Table style White, Table Style Medium 1 on sheet Parts
            var t4 = project.Tasks.Find(t => t.TaskId == "P07-T4");
            Assert.NotNull(t4);
            Assert.Equal(28.4m, t4.MaxScore);
            Assert.Equal(SpecialConditionTypes.ExcelTableCreate, t4.SpecialCondition?.Type);
            Assert.NotNull(t4.SpecialCondition?.ExcelTableCreateConfig);
            Assert.Equal("Parts", t4.SpecialCondition.ExcelTableCreateConfig.WorksheetName);
            Assert.Equal("TableStyleMedium1", t4.SpecialCondition.ExcelTableCreateConfig.ExpectedTableStyle);

            // Assert T5: Line Sparkline at F4 on sheet Parts
            var t5 = project.Tasks.Find(t => t.TaskId == "P07-T5");
            Assert.NotNull(t5);
            Assert.Equal(28.4m, t5.MaxScore);
            Assert.Equal(SpecialConditionTypes.ExcelSparkline, t5.SpecialCondition?.Type);
            Assert.NotNull(t5.SpecialCondition?.ExcelSparklineConfig);
            Assert.Equal("Parts", t5.SpecialCondition.ExcelSparklineConfig.WorksheetName);
            Assert.Equal("line", t5.SpecialCondition.ExcelSparklineConfig.SparklineType);
            Assert.Equal("B4:D4", t5.SpecialCondition.ExcelSparklineConfig.DataRange);
            Assert.Equal("F4", t5.SpecialCondition.ExcelSparklineConfig.LocationRange);

            // Assert Supported Types
            Assert.Contains(SpecialConditionTypes.ExcelSparkline, SpecialConditionTypes.Supported);
            Assert.Contains(SpecialConditionTypes.ExcelTableRowDelete, SpecialConditionTypes.Supported);
        }
    }
}
