using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project10
{
    public class P10T2Grader : ITaskGrader
    {
        public string TaskId => "P10-T2";
        public string TaskName => "Enrollment summary: tao Named Range Enrollment";
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
                decimal score = 0m;
                if (!P10GraderHelpers.TryGetDefinedName(studentSheet.Workbook, "Enrollment", out var definedValue))
                {
                    result.Errors.Add("Không tìm thấy Named Range 'Enrollment'.");
                    return result;
                }

                score += 1m;
                result.Details.Add("Tìm thấy Named Range 'Enrollment'.");

                if (P10GraderHelpers.IsRangeMatch(definedValue, "A3:B7"))
                {
                    score += 2m;
                    result.Details.Add("Named Range 'Enrollment' dung vung A3:B7.");
                }
                else
                {
                    result.Errors.Add($"Named Range 'Enrollment' sai vung. Hiện tại: '{definedValue}'.");
                }

                var normalizedDefinedValue = (definedValue ?? string.Empty)
                    .Replace(" ", string.Empty, StringComparison.Ordinal)
                    .Replace("'", string.Empty, StringComparison.Ordinal)
                    .ToUpperInvariant();
                if (normalizedDefinedValue.Contains("ENROLLMENTSUMMARY!", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Named Range dung sheet 'Enrollment summary'.");
                }
                else
                {
                    result.Errors.Add($"Named Range chưa trỏ dung sheet 'Enrollment summary'. Hiện tại: '{definedValue}'.");
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



