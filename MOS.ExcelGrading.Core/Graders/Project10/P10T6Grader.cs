using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MOS.ExcelGrading.Core.Graders.Project10
{
    public class P10T6Grader : ITaskGrader
    {
        public string TaskId => "P10-T6";
        public string TaskName => "Income: line chart";
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
                var ws = P10GraderHelpers.GetSheet(studentSheet.Workbook, "Income");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Income'.");
                    return result;
                }

                decimal score = 0m;
                var chart = ws.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Khong tim thay chart tren sheet Income.");
                    return result;
                }

                score += 1m;
                result.Details.Add("Tim thay chart tren sheet Income.");

                if (chart.ChartType == eChartType.Line
                    || chart.ChartType == eChartType.LineMarkers
                    || chart.ChartType == eChartType.LineStacked
                    || chart.ChartType == eChartType.LineMarkersStacked)
                {
                    score += 1m;
                    result.Details.Add("Chart dung nhom Line.");
                }
                else
                {
                    result.Errors.Add($"Loai chart chua dung. Hien tai: {chart.ChartType}.");
                }

                if (P10GraderHelpers.IsSeriesRangeMatch(chart, "A4:A7", "B4:B7"))
                {
                    score += 1m;
                    result.Details.Add("Du lieu chart dung: X=A4:A7, Y=B4:B7.");
                }
                else
                {
                    var series = chart.Series.FirstOrDefault();
                    result.Errors.Add(
                        $"Du lieu chart chua dung. Hien tai: X='{series?.XSeries}', Y='{series?.Series}'.");
                }

                if (P10GraderHelpers.IsChartBoundsMatch(chart, "A10:G24"))
                {
                    score += 1m;
                    result.Details.Add("Vi tri chart dung: A10:G24.");
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
