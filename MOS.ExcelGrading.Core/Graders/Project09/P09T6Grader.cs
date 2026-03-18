using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MOS.ExcelGrading.Core.Graders.Project09
{
    public class P09T6Grader : ITaskGrader
    {
        public string TaskId => "P09-T6";
        public string TaskName => "Create 3D Pie Chart in Farmers & Market sheet";
        public decimal MaxScore => 8;

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
                var workbook = studentSheet.Workbook;
                var ws = P09GraderHelpers.GetSheet(workbook, "Farmers & Market")
                         ?? P09GraderHelpers.GetSheet(workbook, "Farmers & Markets")
                         ?? P09GraderHelpers.GetSheet(workbook, "Farmer & Market");

                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Farmers & Market'.");
                    return result;
                }

                decimal score = 0m;
                var pieCharts = ws.Drawings.OfType<ExcelChart>()
                    .Where(chart => chart.ChartType == eChartType.Pie3D || chart.ChartType == eChartType.PieExploded3D)
                    .ToList();

                if (pieCharts.Count == 0)
                {
                    result.Errors.Add("Khong tim thay chart Pie 3D tren sheet Farmers & Market.");
                    return result;
                }

                score += 2m;
                result.Details.Add($"Tim thay {pieCharts.Count} chart Pie 3D.");

                var targetChart = pieCharts
                    .FirstOrDefault(chart => IsTargetDataRange(chart))
                    ?? pieCharts.First();

                if (IsTargetDataRange(targetChart))
                {
                    score += 3m;
                    result.Details.Add("Vung du lieu dung: Category B13:B18 va Value F13:F18.");
                }
                else
                {
                    var firstSeries = targetChart.Series.FirstOrDefault();
                    result.Errors.Add(
                        $"Vung du lieu chua dung. X='{firstSeries?.XSeries}', Y='{firstSeries?.Series}'.");
                }

                var fromRow = targetChart.From.Row + 1;
                var fromCol = targetChart.From.Column + 1;
                var toRow = targetChart.To.Row + 1;
                var toCol = targetChart.To.Column + 1;

                var expectedFromRow = 2;
                var expectedFromCol = 10; // J
                var expectedToRow = 17;
                var expectedToCol = 16; // P

                if (fromRow == expectedFromRow
                    && fromCol == expectedFromCol
                    && toRow == expectedToRow
                    && toCol == expectedToCol)
                {
                    score += 3m;
                    result.Details.Add("Vi tri chart dung vung J2:P17.");
                }
                else
                {
                    var currentBounds =
                        $"{P09GraderHelpers.CellAddress(fromRow, fromCol)}:{P09GraderHelpers.CellAddress(toRow, toCol)}";
                    result.Errors.Add($"Vi tri chart chua dung. Hien tai: {currentBounds}, mong doi: J2:P17.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }

        private static bool IsTargetDataRange(ExcelChart chart)
        {
            var series = chart.Series.FirstOrDefault();
            if (series == null)
            {
                return false;
            }

            var xRange = P09GraderHelpers.NormalizeRange(series.XSeries?.ToString());
            var yRange = P09GraderHelpers.NormalizeRange(series.Series?.ToString());
            return string.Equals(xRange, "B13:B18", StringComparison.OrdinalIgnoreCase)
                   && string.Equals(yRange, "F13:F18", StringComparison.OrdinalIgnoreCase);
        }
    }
}
