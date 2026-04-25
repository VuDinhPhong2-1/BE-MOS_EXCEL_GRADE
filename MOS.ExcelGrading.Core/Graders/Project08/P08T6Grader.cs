using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MOS.ExcelGrading.Core.Graders.Project08
{
    public class P08T6Grader : ITaskGrader
    {
        public string TaskId => "P08-T6";
        public string TaskName => "Summary chart mở rộng để thêm Current Year";
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
                var ws = P08GraderHelpers.GetSheet(studentSheet, "Summary");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Summary'.");
                    return result;
                }

                var chart = ws.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Không tìm thấy chart trên sheet 'Summary'.");
                    return result;
                }

                decimal score = 1m; // Co chart.
                if (chart.Series.Count >= 2)
                {
                    score += 1m;
                    result.Details.Add($"Chart co {chart.Series.Count} series.");
                }
                else
                {
                    result.Errors.Add("Chart chưa mở rộng series (cần ít nhất 2 series).");
                }

                var hasCurrentYearSeries = false;
                var hasPreviousYearSeries = false;

                foreach (var series in chart.Series)
                {
                    var yRange = P08GraderHelpers.NormalizeAddress(series.Series);
                    var header = P08GraderHelpers.NormalizeAddress(series.HeaderAddress?.Address);

                    if (yRange == "C6:C12" || header == "C5")
                    {
                        hasCurrentYearSeries = true;
                    }

                    if (yRange == "B6:B12" || header == "B5")
                    {
                        hasPreviousYearSeries = true;
                    }
                }

                if (hasCurrentYearSeries)
                {
                    score += 1m;
                    result.Details.Add("Đã thêm series Current Year vào chart.");
                }
                else
                {
                    result.Errors.Add("Không tìm thấy series Current Year (C6:C12) trong chart.");
                }

                if (hasPreviousYearSeries)
                {
                    score += 1m;
                    result.Details.Add("Series cũ vẫn được giữ lại sau khi mở rộng.");
                }
                else
                {
                    result.Errors.Add("Không tìm thấy series năm trước (B6:B12) trong chart.");
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

// minor-sync: non-functional graders update
