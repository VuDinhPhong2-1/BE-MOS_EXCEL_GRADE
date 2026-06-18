using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project01
{
    public class P01T3Grader : ITaskGrader
    {
        public string TaskId => "P01-T3";
        public string TaskName => "Tính t?ng t?i C48 b?ng SUM và 4 named range";
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
                var studentMenu = studentSheet.Workbook.Worksheets["Menu Items"];

                if (studentMenu == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'Menu Items'");
                    return result;
                }

                var studentFormula = NormalizeFormula(studentMenu.Cells["C48"].Formula);
                decimal score = 0;

                if (string.IsNullOrWhiteSpace(studentFormula))
                {
                    result.Errors.Add("Ô C48 chua có công th?c");
                    result.Score = 0;
                    return result;
                }

                score += 1;
                result.Details.Add("Có công th?c t?i C48");

                if (studentFormula.Contains("SUM(", StringComparison.OrdinalIgnoreCase))
                {
                    score += 1;
                    result.Details.Add("Công th?c dùng hàm SUM");
                }
                else
                {
                    result.Errors.Add("Công th?c chua dùng hàm SUM");
                }

                // named range d?u có 2 bi?n th? chính t? ph? bi?n: SPECIALITY/SPECIALTY.
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
                    result.Details.Add("Công th?c dùng d? 4 named range yêu c?u");
                }
                else
                {
                    result.Errors.Add($"Công th?c chua dùng d? named range ({hitCount}/4)");
                }

                result.Score = score;
            }
            catch (Exception ex)
            {
                result.Errors.Add($"L?i: {ex.Message}");
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

// minor-sync: non-functional graders update

