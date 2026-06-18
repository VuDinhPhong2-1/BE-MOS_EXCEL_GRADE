using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project01
{
    public class P01T4Grader : ITaskGrader
    {
        public string TaskId => "P01-T4";
        public string TaskName => "Ð?m s? m?c thi?u hàng tháng 9 t?i ô K48";
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

                var studentFormula = NormalizeFormula(studentMenu.Cells["K48"].Formula);

                decimal score = 0;

                if (!string.IsNullOrWhiteSpace(studentFormula))
                {
                    score += 1;
                    result.Details.Add("Có công th?c t?i K48");
                }
                else
                {
                    result.Errors.Add("Ô K48 chua có công th?c");
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
                    result.Details.Add("Có s? d?ng hàm d?m");
                }
                else
                {
                    result.Errors.Add("Công th?c chua dùng hàm d?m phù h?p");
                }

                if (referencesSepColumn)
                {
                    score += 1;
                    result.Details.Add("Công th?c tham chi?u dúng c?t Sep c?n ki?m tra");
                }
                else
                {
                    result.Errors.Add("Công th?c chua tham chi?u c?t Sep (Table1[Sep])");
                }

                if (checksBlank)
                {
                    score += 1;
                    result.Details.Add("Công th?c có di?u ki?n d?m ô tr?ng phù h?p");
                }
                else
                {
                    result.Errors.Add("Thi?u di?u ki?n d?m ô tr?ng");
                }

                result.Score = Math.Min(MaxScore, score);
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

