using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MOS.ExcelGrading.Core.Graders.Project06
{
    public class P06T6Grader : ITaskGrader
    {
        public string TaskId => "P06-T6";
        public string TaskName => "Move pie chart from 'Qtr 2' to chart sheet 'Qtr 2 Chart'";
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
                var wb = studentSheet.Workbook;
                var qtr2 = P06GraderHelpers.GetSheet(studentSheet, "Qtr 2");
                if (qtr2 == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Qtr 2'.");
                    return result;
                }

                decimal score = 0;
                var chartSheet = wb.Worksheets.FirstOrDefault(w =>
                    w is ExcelChartsheet &&
                    string.Equals(w.Name ?? string.Empty, "Qtr 2 Chart", StringComparison.Ordinal));
                if (chartSheet != null)
                {
                    score += 1m;
                    result.Details.Add("Da tao chart sheet 'Qtr 2 Chart'.");
                }
                else
                {
                    result.Errors.Add("Khong tim thay chart sheet 'Qtr 2 Chart' (can dung dung ten).");
                    result.Score = score;
                    return result;
                }

                var movedChart = chartSheet.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (movedChart != null && movedChart.ChartType == eChartType.Pie)
                {
                    score += 1m;
                    result.Details.Add("Chart trong chart sheet co dang Pie.");
                }
                else if (movedChart != null)
                {
                    result.Errors.Add($"Chart trong chart sheet khong phai Pie (hien tai: {movedChart.ChartType}).");
                }
                else
                {
                    result.Errors.Add("Chart sheet chua co chart.");
                }

                var qtr2StillHasChart = qtr2.Drawings.OfType<ExcelChart>().Any();
                if (!qtr2StillHasChart)
                {
                    score += 1m;
                    result.Details.Add("Sheet 'Qtr 2' khong con chart sau khi move.");
                }
                else
                {
                    result.Errors.Add("Sheet 'Qtr 2' van con chart, chua move dung.");
                }

                if (movedChart != null && movedChart.Series.Count > 0)
                {
                    var fromQtr2 = movedChart.Series.Any(s =>
                        (s.Series ?? string.Empty).Contains("Qtr 2", StringComparison.OrdinalIgnoreCase) ||
                        (s.XSeries ?? string.Empty).Contains("Qtr 2", StringComparison.OrdinalIgnoreCase));
                    if (fromQtr2)
                    {
                        score += 1m;
                        result.Details.Add("Chart da move van tham chieu du lieu tu sheet Qtr 2.");
                    }
                    else
                    {
                        result.Errors.Add("Chart trong chart sheet khong tham chieu du lieu Qtr 2.");
                    }
                }
                else
                {
                    result.Errors.Add("Khong du du lieu series de xac nhan chart move.");
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
