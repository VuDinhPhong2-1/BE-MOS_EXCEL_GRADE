using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MOS.ExcelGrading.Core.Graders.Project10
{
    public class P10T4Grader : ITaskGrader
    {
        public string TaskId => "P10-T4";
        public string TaskName => "Next semester: clustered column chart";
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
                var ws = P10GraderHelpers.GetSheet(studentSheet.Workbook, "Next semester");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Next semester'.");
                    return result;
                }

                decimal score = 0m;
                var chart = ws.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Khong tim thay chart tren sheet Next semester.");
                    return result;
                }

                score += 1m;
                result.Details.Add("Tim thay chart tren sheet Next semester.");

                if (chart.ChartType == eChartType.ColumnClustered)
                {
                    score += 1m;
                    result.Details.Add("Chart dung loai ColumnClustered.");
                }
                else
                {
                    result.Errors.Add($"Loai chart chua dung. Hien tai: {chart.ChartType}.");
                }

                if (P10GraderHelpers.IsSeriesRangeMatch(chart, "A4:A21", "E4:E21"))
                {
                    score += 1m;
                    result.Details.Add("Du lieu chart dung: X=A4:A21, Y=E4:E21.");
                }
                else
                {
                    var series = chart.Series.FirstOrDefault();
                    result.Errors.Add(
                        $"Du lieu chart chua dung. Hien tai: X='{series?.XSeries}', Y='{series?.Series}'.");
                }

                if (P10GraderHelpers.IsChartBoundsMatch(chart, "H3:O17"))
                {
                    score += 1m;
                    result.Details.Add("Vi tri chart dung: H3:O17.");
                }
                else
                {
                    result.Errors.Add($"Vi tri chart chua dung. Hien tai: {P10GraderHelpers.GetChartBounds(chart)}.");
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
