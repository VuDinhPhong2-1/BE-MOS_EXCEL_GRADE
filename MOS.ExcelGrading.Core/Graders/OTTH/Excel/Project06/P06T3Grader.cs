using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project06
{
    public class P06T3Grader : ITaskGrader
    {
        public string TaskId => "P06-T3";
        public string TaskName => "Forecasts Quarter 2 = Quarter 1 * Q2_Increase (named range)";
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
                var ws = P06GraderHelpers.GetSheet(studentSheet, "Forecasts");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Forecasts'.");
                    return result;
                }

                var table = ws.Tables.FirstOrDefault();
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy table trên sheet 'Forecasts'.");
                    return result;
                }

                var q1Col = table.Columns.FirstOrDefault(c =>
                    string.Equals((c.Name ?? string.Empty).Trim(), "Quarter 1", StringComparison.OrdinalIgnoreCase));
                var q2Col = table.Columns.FirstOrDefault(c =>
                    string.Equals((c.Name ?? string.Empty).Trim(), "Quarter 2", StringComparison.OrdinalIgnoreCase));

                if (q1Col == null || q2Col == null)
                {
                    result.Errors.Add("Thiếu cột Quarter 1 hoặc Quarter 2 trong sheet 'Forecasts'.");
                    return result;
                }

                var startRow = table.Address.Start.Row + 1;
                var endRow = table.Address.End.Row - (table.ShowTotal ? 1 : 0);
                var q1ColIndex = table.Address.Start.Column + q1Col.Position;
                var q2ColIndex = table.Address.Start.Column + q2Col.Position;
                var rowCount = Math.Max(0, endRow - startRow + 1);

                if (rowCount == 0)
                {
                    result.Errors.Add("Không có dòng dữ liệu để kiểm tra.");
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
                        result.Errors.Add($"Hàng {row}: Quarter 2 chưa có công thức.");
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
                    result.Details.Add("Tất cả dòng Quarter 2 đều có công thức.");
                }
                if (namedRangeRows == rowCount)
                {
                    result.Details.Add("Tất cả công thức Quarter 2 đều sử dụng named range Q2_Increase.");
                }
                else
                {
                    result.Errors.Add($"Công thức chưa dùng named range Q2_Increase ở {rowCount - namedRangeRows} dòng.");
                }
                if (q1RefRows != rowCount)
                {
                    result.Errors.Add($"Công thức chưa tham chiếu Quarter 1 ở {rowCount - q1RefRows} dòng.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi: {ex.Message}");
            }

            return result;
        }
    }
}

// minor-sync: non-functional graders update

