using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project06
{
    public class P06T3Grader : ITaskGrader
    {
        public string TaskId => "P06-T3";
        public string TaskName => "Forecasts Quarter 2 = Quarter 1 * Q2_Increase (named range)";
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
                var ws = P06GraderHelpers.GetSheet(studentSheet, "Forecasts");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Forecasts'.");
                    return result;
                }

                var table = ws.Tables.FirstOrDefault();
                if (table == null)
                {
                    result.Errors.Add("Khong tim thay table tren Forecasts.");
                    return result;
                }

                var q1Col = table.Columns.FirstOrDefault(c =>
                    string.Equals((c.Name ?? string.Empty).Trim(), "Quarter 1", StringComparison.OrdinalIgnoreCase));
                var q2Col = table.Columns.FirstOrDefault(c =>
                    string.Equals((c.Name ?? string.Empty).Trim(), "Quarter 2", StringComparison.OrdinalIgnoreCase));

                if (q1Col == null || q2Col == null)
                {
                    result.Errors.Add("Thieu cot Quarter 1 hoac Quarter 2 trong Forecasts.");
                    return result;
                }

                var startRow = table.Address.Start.Row + 1;
                var endRow = table.Address.End.Row - (table.ShowTotal ? 1 : 0);
                var q1ColIndex = table.Address.Start.Column + q1Col.Position;
                var q2ColIndex = table.Address.Start.Column + q2Col.Position;
                var rowCount = Math.Max(0, endRow - startRow + 1);

                if (rowCount == 0)
                {
                    result.Errors.Add("Khong co dong du lieu de cham.");
                    return result;
                }

                var formulaRows = 0;
                var namedRangeRows = 0;
                var q1RefRows = 0;

                for (var row = startRow; row <= endRow; row++)
                {
                    var cell = ws.Cells[row, q2ColIndex];
                    var formula = cell.Formula;
                    var normalized = P06GraderHelpers.NormalizeFormula(formula);
                    if (string.IsNullOrWhiteSpace(normalized))
                    {
                        result.Errors.Add($"Hang {row}: Quarter 2 chua co cong thuc.");
                        continue;
                    }

                    formulaRows++;
                    if (normalized.Contains("Q2_INCREASE", StringComparison.Ordinal))
                    {
                        namedRangeRows++;
                    }

                    var q1CellAddress = $"B{row}";
                    var hasQ1Ref =
                        normalized.Contains("QUARTER1", StringComparison.Ordinal) ||
                        normalized.Contains(q1CellAddress, StringComparison.OrdinalIgnoreCase) ||
                        normalized.Contains($"R{row}C{q1ColIndex}", StringComparison.OrdinalIgnoreCase);
                    if (hasQ1Ref)
                    {
                        q1RefRows++;
                    }
                }

                decimal score = 0;
                score += Math.Round(1m * formulaRows / rowCount, 2);
                score += Math.Round(2m * namedRangeRows / rowCount, 2);
                score += Math.Round(1m * q1RefRows / rowCount, 2);

                if (formulaRows == rowCount)
                {
                    result.Details.Add("Tat ca dong Quarter 2 da co cong thuc.");
                }
                if (namedRangeRows == rowCount)
                {
                    result.Details.Add("Tat ca cong thuc Quarter 2 da dung named range Q2_Increase.");
                }
                else
                {
                    result.Errors.Add($"Cong thuc chua dung named range Q2_Increase o {rowCount - namedRangeRows} dong.");
                }
                if (q1RefRows != rowCount)
                {
                    result.Errors.Add($"Cong thuc chua tham chieu Quarter 1 o {rowCount - q1RefRows} dong.");
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
