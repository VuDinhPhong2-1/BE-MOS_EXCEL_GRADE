using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project10
{
    public class P10T5Grader : ITaskGrader
    {
        public string TaskId => "P10-T5";
        public string TaskName => "Next semester: tao bieu do Clustered Column Program va Average cost";
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
                    result.Errors.Add("Không tìm thấy sheet 'Next semester'.");
                    return result;
                }

                decimal score = 0m;
                var chart = ws.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Không tìm thấy chart tren sheet 'Next semester'.");
                    return result;
                }

                score += 1m;
                result.Details.Add("Tìm thấy chart tren sheet 'Next semester'.");

                if (chart.ChartType == eChartType.ColumnClustered)
                {
                    score += 1m;
                    result.Details.Add("Chart dung loai Clustered Column.");
                }
                else
                {
                    result.Errors.Add($"Loai chart chưa đúng. Hiện tại: {chart.ChartType}.");
                }

                if (P10GraderHelpers.IsSeriesRangeMatch(chart, "A4:A21", "E4:E21"))
                {
                    score += 1m;
                    result.Details.Add("Series chart dung Program (A4:A21) va Average cost (E4:E21).");
                }
                else
                {
                    var series = chart.Series.FirstOrDefault();
                    result.Errors.Add($"Series chart chưa đúng. X='{series?.XSeries}', Y='{series?.Series}'.");
                }

                if (P10GraderHelpers.IsChartWithinBounds(chart, "G1:Q21"))
                {
                    score += 1m;
                    result.Details.Add("Vi tri chart hop le o bên phải bang.");
                }
                else
                {
                    result.Errors.Add($"Vi tri chart chưa đúng. Hiện tại: {P10GraderHelpers.GetChartBounds(chart)}.");
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



