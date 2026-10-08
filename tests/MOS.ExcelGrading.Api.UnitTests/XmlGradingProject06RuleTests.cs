using System;
using System.IO;
using System.Text.Json;
using MOS.ExcelGrading.Core.Models;
using Xunit;

namespace MOS.ExcelGrading.Api.UnitTests
{
    public class XmlGradingProject06RuleTests
    {
        [Fact]
        public void Project06Rules_Json_LoadsCorrectly_AndPassesAllValidations()
        {
            var possiblePaths = new[]
            {
                Path.Combine(AppContext.BaseDirectory, "project06.json"),
                Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "tools", "Project06RuleSeeder", "project06.json"),
                Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "tools", "Project06RuleSeeder", "project06.json"),
                Path.Combine(Directory.GetCurrentDirectory(), "tools", "Project06RuleSeeder", "project06.json"),
                Path.Combine(Directory.GetCurrentDirectory(), "BACKEND", "tools", "Project06RuleSeeder", "project06.json")
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
            Assert.Equal("project06", project.ProjectCode);
            Assert.Equal(142.0m, project.MaxScore);
            Assert.Equal(5, project.Tasks.Count);

            var t1 = project.Tasks.Find(t => t.TaskId == "P06-T1");
            Assert.NotNull(t1);
            Assert.Equal(28.4m, t1.MaxScore);
            Assert.Equal(SpecialConditionTypes.ExcelTableTotalRow, t1.SpecialCondition?.Type);
            Assert.NotNull(t1.SpecialCondition?.ExcelTableTotalRowConfig);
            Assert.Equal("Qtr 1", t1.SpecialCondition.ExcelTableTotalRowConfig.WorksheetName);
            Assert.True(t1.SpecialCondition.ExcelTableTotalRowConfig.RequireTotalRow);
            Assert.Equal("Entries", t1.SpecialCondition.ExcelTableTotalRowConfig.ColumnName);

            var t2 = project.Tasks.Find(t => t.TaskId == "P06-T2");
            Assert.NotNull(t2);
            Assert.Equal(28.4m, t2.MaxScore);
            Assert.Equal(SpecialConditionTypes.ExcelFormulaReferences, t2.SpecialCondition?.Type);
            Assert.NotNull(t2.SpecialCondition?.ExcelFormulaReferencesConfig);
            Assert.Equal("Qtr 1", t2.SpecialCondition.ExcelFormulaReferencesConfig.WorksheetName);
            Assert.Equal("E10,E11,E12,E13,E14", t2.SpecialCondition.ExcelFormulaReferencesConfig.Cell);
            Assert.Contains("MAX", t2.SpecialCondition.ExcelFormulaReferencesConfig.RequiredFunctions);

            var t3 = project.Tasks.Find(t => t.TaskId == "P06-T3");
            Assert.NotNull(t3);
            Assert.Equal(28.4m, t3.MaxScore);
            Assert.Equal(SpecialConditionTypes.ExcelChartType, t3.SpecialCondition?.Type);
            Assert.NotNull(t3.SpecialCondition?.ExcelChartTypeConfig);
            Assert.Equal("Qtr 1", t3.SpecialCondition.ExcelChartTypeConfig.WorksheetName);
            Assert.Equal("2-D Pie", t3.SpecialCondition.ExcelChartTypeConfig.ExpectedChartType);
            Assert.Contains("Entries", t3.SpecialCondition.ExcelChartTypeConfig.RequiredFormulaFragments);
            Assert.Contains("Total", t3.SpecialCondition.ExcelChartTypeConfig.RequiredFormulaFragments);
            Assert.Equal(11, t3.SpecialCondition.ExcelChartTypeConfig.PlacedBelowRow);

            var t4 = project.Tasks.Find(t => t.TaskId == "P06-T4");
            Assert.NotNull(t4);
            Assert.Equal(28.4m, t4.MaxScore);
            Assert.Equal(SpecialConditionTypes.ExcelTableCreate, t4.SpecialCondition?.Type);
            Assert.NotNull(t4.SpecialCondition?.ExcelTableCreateConfig);
            Assert.Equal("Qtr 2", t4.SpecialCondition.ExcelTableCreateConfig.WorksheetName);
            Assert.Equal("A2:E10", t4.SpecialCondition.ExcelTableCreateConfig.ExpectedRange);
            Assert.True(t4.SpecialCondition.ExcelTableCreateConfig.HasHeaderRow);
            Assert.Equal("TableStyleLight14", t4.SpecialCondition.ExcelTableCreateConfig.ExpectedTableStyle);

            var t5 = project.Tasks.Find(t => t.TaskId == "P06-T5");
            Assert.NotNull(t5);
            Assert.Equal(28.4m, t5.MaxScore);
            Assert.Equal(SpecialConditionTypes.ExcelChartQuickLayout, t5.SpecialCondition?.Type);
            Assert.NotNull(t5.SpecialCondition?.ExcelChartQuickLayoutConfig);
            Assert.Equal("Qtr 1", t5.SpecialCondition.ExcelChartQuickLayoutConfig.WorksheetName);
            Assert.Equal(2, t5.SpecialCondition.ExcelChartQuickLayoutConfig.LayoutNumber);
            Assert.True(t5.SpecialCondition.ExcelChartQuickLayoutConfig.RequireDataLabels);

            Assert.Contains(SpecialConditionTypes.ExcelTableTotalRow, SpecialConditionTypes.Supported);
            Assert.Contains(SpecialConditionTypes.ExcelTableCreate, SpecialConditionTypes.Supported);
            Assert.Contains(SpecialConditionTypes.ExcelChartQuickLayout, SpecialConditionTypes.Supported);
            Assert.Contains(SpecialConditionTypes.ExcelChartType, SpecialConditionTypes.Supported);
            Assert.Contains(SpecialConditionTypes.ExcelFormulaReferences, SpecialConditionTypes.Supported);
        }
    }
}