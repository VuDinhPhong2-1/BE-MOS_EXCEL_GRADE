using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MOS.ExcelGrading.Core.Graders.Project10
{
    public class P10T1Grader : ITaskGrader
    {
        public string TaskId => "P10-T1";
        public string TaskName => "Enrollment summary: 3D Pie chart";
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
                var ws = P10GraderHelpers.GetSheet(studentSheet.Workbook, "Enrollment summary");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Enrollment summary'.");
                    return result;
                }

                decimal score = 0m;
                var chart = ws.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Khong tim thay chart tren sheet Enrollment summary.");
                    return result;
                }

                if (chart.ChartType == eChartType.Pie3D || chart.ChartType == eChartType.PieExploded3D)
                {
                    score += 1m;
                    result.Details.Add("Chart dung loai Pie 3D.");
                }
                else
                {
                    result.Errors.Add($"Loai chart chua dung. Hien tai: {chart.ChartType}.");
                }

                if (P10GraderHelpers.IsChartBoundsMatch(chart, "A10:G24"))
                {
                    score += 1m;
                    result.Details.Add("Vi tri chart dung vung A10:G24.");
                }
                else
                {
                    result.Errors.Add($"Vi tri chart chua dung. Hien tai: {P10GraderHelpers.GetChartBounds(chart)}.");
                }

                var series = chart.Series.FirstOrDefault();
                if (series != null
                    && P10GraderHelpers.IsRangeMatch(series.XSeries?.ToString(), "A4:A7")
                    && P10GraderHelpers.IsRangeMatch(series.Series?.ToString(), "B4:B7"))
                {
                    score += 2m;
                    result.Details.Add("Du lieu chart dung: X=A4:A7, Y=B4:B7.");
                }
                else
                {
                    result.Errors.Add("Du lieu chart chua dung (can X=A4:A7, Y=B4:B7).");
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
