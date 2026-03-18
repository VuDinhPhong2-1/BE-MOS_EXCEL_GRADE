using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project11
{
    public class P11T2Grader : ITaskGrader
    {
        public string TaskId => "P11-T2";
        public string TaskName => "Shareholders Info: hyperlink in C5";
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
                var ws = P11GraderHelpers.GetSheet(studentSheet.Workbook, "Shareholders Info");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Shareholders Info'.");
                    return result;
                }

                decimal score = 0m;
                var hyperlink = ws.Cells["C5"].Hyperlink;
                if (hyperlink != null)
                {
                    score += 1m;
                    result.Details.Add("Cell C5 da co hyperlink.");
                }
                else
                {
                    result.Errors.Add("Cell C5 chua co hyperlink.");
                    result.Score = score;
                    return result;
                }

                var linkText = hyperlink.OriginalString ?? hyperlink.ToString() ?? string.Empty;
                var normalizedLink = P11GraderHelpers.NormalizeUrl(linkText);
                if (normalizedLink.Contains("tailspintoys.com/beyond.html", StringComparison.Ordinal))
                {
                    score += 2m;
                    result.Details.Add("Hyperlink C5 dung URL yeu cau.");
                }
                else
                {
                    result.Errors.Add($"URL C5 chua dung. Hien tai: '{linkText}'.");
                }

                var cellText = ws.Cells["C5"].Text ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(cellText)
                    && cellText.Contains("Beyond the Depths", StringComparison.OrdinalIgnoreCase))
                {
                    score += 1m;
                    result.Details.Add("Noi dung hien thi o C5 hop le.");
                }
                else
                {
                    result.Errors.Add($"Noi dung C5 chua dung. Hien tai: '{cellText}'.");
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
