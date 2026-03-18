using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project10
{
    public class P10T5Grader : ITaskGrader
    {
        public string TaskId => "P10-T5";
        public string TaskName => "Income: table range and style";
        public decimal MaxScore => 4;

        public TaskResult Grade(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var ws = P10GraderHelpers.GetSheet(studentSheet.Workbook, "Income");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Income'.");
                    return result;
                }

                decimal score = 0m;
                var table = P10GraderHelpers.FindTableByAddress(ws, "A3:B7");
                if (table != null)
                {
                    score += 2m;
                    result.Details.Add("Table dung range A3:B7.");
                }
                else
                {
                    result.Errors.Add($"Khong tim thay table A3:B7. Hien tai: {P10GraderHelpers.JoinTableAddresses(ws)}");
                }

                if (table != null
                    && table.TableStyle == TableStyles.Light14)
                {
                    score += 1m;
                    result.Details.Add("Table style dung: TableStyleLight14.");
                }
                else if (table != null)
                {
                    result.Errors.Add($"Table style chua dung. Hien tai: {table.TableStyle}.");
                }

                if (table != null
                    && table.Columns.Count == 2
                    && string.Equals(table.Columns[0].Name, "Program", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(table.Columns[1].Name, "Total", StringComparison.OrdinalIgnoreCase))
                {
                    score += 1m;
                    result.Details.Add("Cot table dung: Program, Total.");
                }
                else if (table != null)
                {
                    var columnNames = string.Join(", ", table.Columns.Select(c => c.Name));
                    result.Errors.Add($"Ten cot table chua dung. Hien tai: {columnNames}.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }
}
