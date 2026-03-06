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

                var usesCountBlank = studentFormula.Contains("COUNTBLANK(", StringComparison.OrdinalIgnoreCase);
                var usesCountIfFamily = studentFormula.Contains("COUNTIF(", StringComparison.OrdinalIgnoreCase)
                    || studentFormula.Contains("COUNTIFS(", StringComparison.OrdinalIgnoreCase);
                var referencesSepColumn = studentFormula.Contains("TABLE1[SEP]", StringComparison.OrdinalIgnoreCase)
                    || studentFormula.Contains("[SEP]", StringComparison.OrdinalIgnoreCase)
                    || studentFormula.Contains("K22:K46", StringComparison.OrdinalIgnoreCase);
                var checksBlank = usesCountBlank
                    || studentFormula.Contains("\"\"");

                if (usesCountBlank || usesCountIfFamily)
                {
                    score += 1;
                    result.Details.Add("Có sử dụng hàm đếm");
                }
                else
                {
                    result.Errors.Add("Công thức chưa dùng hàm đếm phù hợp");
                }

                if (referencesSepColumn)
                {
                    score += 1;
                    result.Details.Add("Công thức tham chiếu đúng cột Sep cần kiểm tra");
                }
                else
                {
                    result.Errors.Add("Công thức chưa tham chiếu cột Sep (Table1[Sep])");
                }

                if (checksBlank)
                {
                    score += 1;
                    result.Details.Add("Công thức có điều kiện đếm ô trống phù hợp");
                }
                else
                {
                    result.Errors.Add("Thiếu điều kiện đếm ô trống");
                }

                result.Score = Math.Min(MaxScore, score);
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
