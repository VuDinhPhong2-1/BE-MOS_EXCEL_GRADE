using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project20
{
    public class P20T4Grader : ITaskGrader
    {
        public string TaskId => "P20-T4";
        public string TaskName => "Trong trang tính “New York City”, tại ô D23, sử dụng hàm để hiển thị giá trị lớn nhất trong cột Air Miles.";
        public decimal MaxScore => 18m;

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
                var worksheet = P20GraderHelpers.GetSheet(studentSheet.Workbook, "New York City");
                if (worksheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'New York City'.");
                    return result;
                }

                var table = P20GraderHelpers.FindTable(worksheet, "Table14", "Air Miles");
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy bảng Table14 hoặc cột 'Air Miles' để kiểm tra công thức tại D23.");
                    return result;
                }

                if (!P20GraderHelpers.TryGetColumnIndex(table, "Air Miles", out var airMilesCol))
                {
                    result.Errors.Add("Không xác định được cột 'Air Miles' trong bảng Table14.");
                    return result;
                }

                decimal score = 0m;
                var formulaCell = worksheet.Cells["D23"];
                var formula = formulaCell.Formula ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(formula))
                {
                    score += 4m;
                    result.Details.Add("Ô D23 đã có công thức.");
                }
                else
                {
                    result.Errors.Add("Ô D23 chưa có công thức.");
                }

                var dataStartRow = table.Address.Start.Row + 1;
                var dataEndRow = table.Address.End.Row;
                var maxFormulaOk = P20GraderHelpers.FormulaLooksLikeMaxOverAirMiles(formula, airMilesCol, dataStartRow, dataEndRow);
                if (maxFormulaOk)
                {
                    score += 8m;
                    result.Details.Add("Công thức D23 dùng MAX và tham chiếu đúng phạm vi cột Air Miles (không lấy dư phạm vi).");
                }
                else
                {
                    result.Errors.Add($"Công thức D23 chưa đúng cấu trúc MAX cho cột Air Miles. Hiện tại: '{formula}'.");
                }

                decimal expectedMax = decimal.MinValue;
                var hasAnyValue = false;
                for (var row = dataStartRow; row <= dataEndRow; row++)
                {
                    if (P20GraderHelpers.TryGetNumericValue(worksheet.Cells[row, airMilesCol], out var value))
                    {
                        expectedMax = Math.Max(expectedMax, value);
                        hasAnyValue = true;
                    }
                }

                if (!hasAnyValue)
                {
                    result.Errors.Add("Không đọc được dữ liệu số trong cột Air Miles để đối chiếu kết quả D23.");
                    result.Score = Math.Min(MaxScore, score);
                    return result;
                }

                if (P20GraderHelpers.TryGetNumericValue(formulaCell, out var actualMax)
                    && Math.Abs(actualMax - expectedMax) <= 0.01m)
                {
                    score += 6m;
                    result.Details.Add("Giá trị kết quả tại D23 đúng bằng giá trị lớn nhất của cột Air Miles.");
                }
                else
                {
                    result.Errors.Add(
                        $"Giá trị tại D23 chưa đúng. Hiện tại: '{formulaCell.Text}', mong đợi: '{expectedMax}'.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 4: {ex.Message}.");
            }

            return result;
        }
    }
}

