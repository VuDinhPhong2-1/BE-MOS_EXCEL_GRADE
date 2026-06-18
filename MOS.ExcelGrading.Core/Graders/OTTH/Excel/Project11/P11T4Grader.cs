using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project11
{
    public class P11T4Grader : ITaskGrader
    {
        public string TaskId => "P11-T4";
        public string TaskName => "Shareholders Info: thiet lap Hyperlink va Display Text cho C5";
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
                var ws = P11GraderHelpers.GetSheet(studentSheet.Workbook, "Shareholders Info");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'Shareholders Info'.");
                    return result;
                }

                decimal score = 0m;
                var hyperlink = ws.Cells["C5"].Hyperlink;
                if (hyperlink != null)
                {
                    score += 1m;
                    result.Details.Add("C5 da co hyperlink.");
                }
                else
                {
                    result.Errors.Add("C5 chua có hyperlink.");
                    result.Score = score;
                    return result;
                }

                var linkText = hyperlink.OriginalString ?? hyperlink.ToString() ?? string.Empty;
                var normalized = P11GraderHelpers.NormalizeUrl(linkText);
                if (string.Equals(normalized, "http://tailspintoys.com/beyond.html", StringComparison.Ordinal))
                {
                    score += 2m;
                    result.Details.Add("Hyperlink dung URL yêu c?u.");
                }
                else
                {
                    result.Errors.Add($"URL hyperlink chua dúng. Hi?n t?i: '{linkText}'.");
                }

                var displayText = ws.Cells["C5"].Text ?? string.Empty;
                if (string.Equals(displayText, "More Info", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Van ban hien thi o C5 dung chinh ta: 'More Info'.");
                }
                else
                {
                    result.Errors.Add($"Van ban hien thi C5 chua dúng. Hi?n t?i: '{displayText}'.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"L?i: {ex.Message}");
            }

            return result;
        }
    }
}




