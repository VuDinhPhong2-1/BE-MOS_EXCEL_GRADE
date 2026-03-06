using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Project01
{
    public class P01T3Grader : ITaskGrader
    {
        public string TaskId => "P01-T3";
        public string TaskName => "Tính tổng tại C48 bằng SUM và 4 named range";
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

                var studentFormula = NormalizeFormula(studentMenu.Cells["C48"].Formula);
                decimal score = 0;

                if (string.IsNullOrWhiteSpace(studentFormula))
                {
                    result.Errors.Add("Ô C48 chưa có công thức");
                    result.Score = 0;
                    return result;
                }

                score += 1;
                result.Details.Add("Có công thức tại C48");

                if (studentFormula.Contains("SUM(", StringComparison.OrdinalIgnoreCase))
                {
                    score += 1;
                    result.Details.Add("Công thức dùng hàm SUM");
                }
                else
                {
                    result.Errors.Add("Công thức chưa dùng hàm SUM");
                }

                // named range đầu có 2 biến thể chính tả phổ biến: SPECIALITY/SPECIALTY.
                var hitCount = 0;
                if (studentFormula.Contains("SPECIALITY_TOTAL", StringComparison.OrdinalIgnoreCase)
                    || studentFormula.Contains("SPECIALTY_TOTAL", StringComparison.OrdinalIgnoreCase))
                {
                    hitCount++;
                }

                if (studentFormula.Contains("SMOOTHIES_TOTAL", StringComparison.OrdinalIgnoreCase))
                {
                    hitCount++;
                }

                if (studentFormula.Contains("SANDWICHES_TOTAL", StringComparison.OrdinalIgnoreCase))
                {
                    hitCount++;
                }

                if (studentFormula.Contains("SOUPS_TOTAL", StringComparison.OrdinalIgnoreCase))
                {
                    hitCount++;
                }

                if (hitCount == 4)
                {
                    score += 2;
                    result.Details.Add("Công thức dùng đủ 4 named range yêu cầu");
                }
                else
                {
                    result.Errors.Add($"Công thức chưa dùng đủ named range ({hitCount}/4)");
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
