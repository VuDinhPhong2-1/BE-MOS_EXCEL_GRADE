using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project10
{
    public class P10T6Grader : ITaskGrader
    {
        public string TaskId => "P10-T6";
        public string TaskName => "Enrollment summary: bieu do Style 7 + Monochromatic Palette 6";
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
                var ws = P10GraderHelpers.GetSheet(studentSheet.Workbook, "Enrollment summary");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Enrollment summary'.");
                    return result;
                }

                decimal score = 0m;
                var chart = ws.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Không tìm thấy chart tren sheet 'Enrollment summary'.");
                    return result;
                }

                score += 1m;
                result.Details.Add("Tìm thấy chart tren sheet 'Enrollment summary'.");

                var (choiceStyle, fallbackStyle) = P10GraderHelpers.GetChartAlternateStyles(chart);
                var styleMatch = string.Equals(choiceStyle, "108", StringComparison.Ordinal)
                                 && string.Equals(fallbackStyle, "8", StringComparison.Ordinal);
                if (styleMatch)
                {
                    score += 1m;
                    result.Details.Add("Chart style XML hop le (Choice=108, Fallback=8), tuong ung Style 7.");
                }
                else
                {
                    result.Errors.Add($"Chart style XML chưa đúng. Choice='{choiceStyle}', Fallback='{fallbackStyle}'.");
                }

                var (colorStyleId, chartStyleId) = P10GraderHelpers.GetStyleManagerIds(chart);
                if (string.Equals(colorStyleId, "19", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Color style ID = 19 (Monochromatic Palette 6).");
                }
                else
                {
                    result.Errors.Add($"Color style ID chưa đúng. Hiện tại: '{colorStyleId}' (mong đợi: 19).");
                }

                if (string.Equals(chartStyleId, "268", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Chart style ID = 268 (Style 7 cho dang chart nay).");
                }
                else
                {
                    result.Errors.Add($"Chart style ID chưa đúng. Hiện tại: '{chartStyleId}' (mong đợi: 268).");
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



