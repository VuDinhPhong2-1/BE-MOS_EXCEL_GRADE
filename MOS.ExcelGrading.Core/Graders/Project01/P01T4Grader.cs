using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Project01
{
    public class P01T4Grader : ITaskGrader
    {
        public string TaskId => "P01-T4";
        public string TaskName => "Đếm số mục thiếu hàng tháng 9 tại ô K48";
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
                var studentMenu = studentSheet.Workbook.Worksheets["Menu Items"];

                if (studentMenu == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Menu Items'");
                    return result;
                }

                var studentFormula = NormalizeFormula(studentMenu.Cells["K48"].Formula);

                decimal score = 0;

                if (!string.IsNullOrWhiteSpace(studentFormula))
                {
                    score += 1;
                    result.Details.Add("Có công thức tại K48");
                }
                else
                {
                    result.Errors.Add("Ô K48 chưa có công thức");
                    result.Score = score;
                    return result;
                }

                if (studentFormula.Contains("COUNT", StringComparison.OrdinalIgnoreCase))
                {
                    score += 1;
                    result.Details.Add("Có sử dụng hàm đếm");
                }
                else
                {
                    result.Errors.Add("Công thức chưa dùng hàm đếm phù hợp");
                }

                if (studentFormula.Contains("COUNTIF(", StringComparison.OrdinalIgnoreCase)
                    || studentFormula.Contains("COUNTIFS(", StringComparison.OrdinalIgnoreCase)
                    || studentFormula.Contains("COUNTBLANK(", StringComparison.OrdinalIgnoreCase))
                {
                    score += 1;
                    result.Details.Add("Dùng đúng nhóm hàm cho bài đếm thiếu hàng");
                }
                else
                {
                    result.Errors.Add("Chưa dùng COUNTIF/COUNTIFS/COUNTBLANK");
                }

                if (studentFormula.Contains("\"\""))
                {
                    score += 1;
                    result.Details.Add("Có điều kiện kiểm tra ô trống (mục thiếu hàng)");
                }
                else
                {
                    result.Errors.Add("Thiếu điều kiện kiểm tra ô trống trong công thức");
                }

                result.Score = score;
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi: {ex.Message}");
            }

            return result;
        }

        private static string NormalizeFormula(string? formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return string.Empty;

            return formula
                .Replace("=", string.Empty)
                .Replace("$", string.Empty)
                .Replace(" ", string.Empty)
                .ToUpperInvariant();
        }
    }
}
