using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project10
{
    public class P10T3Grader : ITaskGrader
    {
        public string TaskId => "P10-T3";
        public string TaskName => "Income: tao Table A3:B7, style Light14";
        public decimal MaxScore => 4;

        public TaskResult Grade(ExcelWorksheet studentSheet)
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
                    result.Errors.Add("Không tìm th?y sheet 'Income'.");
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
                    result.Errors.Add($"Không tìm th?y table A3:B7. Hi?n t?i: {P10GraderHelpers.JoinTableAddresses(ws)}.");
                    result.Score = score;
                    return result;
                }

                if (table.TableStyle == TableStyles.Light14)
                {
                    score += 1m;
                    result.Details.Add("Table style dung: TableStyleLight14.");
                }
                else
                {
                    result.Errors.Add($"Table style chua dúng. Hi?n t?i: {table.TableStyle}.");
                }

                var expectedHeaders = table.Columns.Count == 2
                                      && string.Equals(P10GraderHelpers.NormalizeText(table.Columns[0].Name), "Program", StringComparison.Ordinal)
                                      && string.Equals(P10GraderHelpers.NormalizeText(table.Columns[1].Name), "Total", StringComparison.Ordinal);
                if (expectedHeaders)
                {
                    score += 1m;
                    result.Details.Add("Ten c?t table dung: Program, Total.");
                }
                else
                {
                    var columns = string.Join(", ", table.Columns.Select(c => c.Name));
                    result.Errors.Add($"Ten c?t table chua dúng. Hi?n t?i: {columns}.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"L?i: {ex.Message}");
            }

            return result;
        }
    }
}




