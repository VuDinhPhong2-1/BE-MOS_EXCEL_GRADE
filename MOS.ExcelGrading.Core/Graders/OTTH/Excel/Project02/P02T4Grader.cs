using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project02
{
    public class P02T4Grader : ITaskGrader
    {
        public string TaskId => "P02-T4";
        public string TaskName => "Đếm số tháng không có policy mới trong cột Inactive months";
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
                var ws = studentSheet.Workbook.Worksheets["New Policy"];
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'New Policy'");
                    return result;
                }

                var table = ws.Tables.FirstOrDefault();
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy bảng dữ liệu trên sheet 'New Policy'");
                    return result;
                }

                var inactiveCol = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name?.Trim(), "Inactive months", StringComparison.OrdinalIgnoreCase));
                if (inactiveCol == null)
                {
                    result.Errors.Add("Không tìm thấy cột 'Inactive months'");
                    return result;
                }

                var startRow = table.Address.Start.Row + 1;
                var endRow = table.Address.End.Row - (table.ShowTotal ? 1 : 0);
                var colIndex = table.Address.Start.Column + inactiveCol.Position;
                var totalRows = Math.Max(0, endRow - startRow + 1);

                if (totalRows == 0)
                {
                    result.Errors.Add("Không có dòng dữ liệu để chấm");
                    return result;
                }

                var hasFormulaRows = 0;
                var countFunctionRows = 0;
                var janJunRefRows = 0;
                var blankConditionRows = 0;

                for (var row = startRow; row <= endRow; row++)
                {
                    var formula = NormalizeFormula(ws.Cells[row, colIndex].Formula);
                    if (string.IsNullOrWhiteSpace(formula))
                    {
                        continue;
                    }

                    hasFormulaRows++;
                    if (formula.Contains("COUNTBLANK(") || formula.Contains("COUNTIF(") || formula.Contains("COUNTIFS("))
                    {
                        countFunctionRows++;
                    }

                    if (formula.Contains("JANUARY", StringComparison.OrdinalIgnoreCase) &&
                        formula.Contains("JUNE", StringComparison.OrdinalIgnoreCase))
                    {
                        janJunRefRows++;
                    }

                    if (formula.Contains("COUNTBLANK(") || formula.Contains("\"\""))
                    {
                        blankConditionRows++;
                    }
                }

                decimal score = 0;
                if (hasFormulaRows == totalRows)
                {
                    score += 1m;
                    result.Details.Add("Cột Inactive months đã có công thức cho tất cả dòng");
                }
                else
                {
                    result.Errors.Add($"Thiếu công thức tại cột Inactive months ({hasFormulaRows}/{totalRows})");
                }

                if (countFunctionRows == totalRows)
                {
                    score += 1m;
                    result.Details.Add("Công thức có sử dụng hàm đếm");
                }
                else
                {
                    result.Errors.Add($"Hàm đếm chưa dùng/đủ ({countFunctionRows}/{totalRows})");
                }

                if (janJunRefRows == totalRows)
                {
                    score += 1m;
                    result.Details.Add("Công thức tham chiếu đúng khoảng January:June");
                }
                else
                {
                    result.Errors.Add($"Tham chiếu January:June chưa đúng/đủ ({janJunRefRows}/{totalRows})");
                }

                if (blankConditionRows == totalRows)
                {
                    score += 1m;
                    result.Details.Add("Công thức có điều kiện đếm tháng trong");
                }
                else
                {
                    result.Errors.Add($"Điều kiện đếm ô trống chưa đầy đủ ({blankConditionRows}/{totalRows})");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }

        private static string NormalizeFormula(string? formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                return string.Empty;
            }

            return formula
                .Replace("=", string.Empty)
                .Replace("$", string.Empty)
                .Replace(" ", string.Empty)
                .ToUpperInvariant();
        }
    }
}


// minor-sync: non-functional graders update

