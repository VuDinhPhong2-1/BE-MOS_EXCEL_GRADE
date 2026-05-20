using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MOS.ExcelGrading.Core.Graders.Project06
{
    public class P06T5Grader : ITaskGrader
    {
        public string TaskId => "P06-T5";
        public string TaskName => "Comparison chart: Switch Row/Column";
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
                var sheet = P06GraderHelpers.GetSheet(studentSheet, "Comparison");
                if (sheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Comparison'.");
                    return result;
                }

                var chart = sheet.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Không tìm thấy chart trên sheet Comparison.");
                    return result;
                }

                decimal score = 1m; // Tim thay chart.
                var expectedSeriesCount = 3;
                if (chart.Series.Count == expectedSeriesCount)
                {
                    score += 1m;
                    result.Details.Add("Số series chart = 3 (đúng sau khi Switch Row/Column).");
                }
                else
                {
                    result.Errors.Add($"Số series chart chưa đúng. Hiện tại: {chart.Series.Count}, mong đợi: 3.");
                }

                var xRangeOkCount = 0;
                var yRangeOkCount = 0;
                var headerOkCount = 0;
                for (var i = 0; i < chart.Series.Count; i++)
                {
                    var series = chart.Series[i];
                    var expectedX = "B3:E3";
                    var expectedY = $"B{7 + i}:E{7 + i}";
                    var expectedHeader = $"A{7 + i}";

                    var xRange = P06GraderHelpers.NormalizeAddress(series.XSeries);
                    var yRange = P06GraderHelpers.NormalizeAddress(series.Series);
                    var header = P06GraderHelpers.NormalizeAddress(series.HeaderAddress?.Address);

                    if (xRange == expectedX)
                    {
                        xRangeOkCount++;
                    }
                    if (yRange == expectedY)
                    {
                        yRangeOkCount++;
                    }
                    if (header == expectedHeader)
                    {
                        headerOkCount++;
                    }
                }

                if (chart.Series.Count > 0)
                {
                    score += Math.Round(1m * xRangeOkCount / chart.Series.Count, 2);
                    score += Math.Round(1m * Math.Min(yRangeOkCount, headerOkCount) / chart.Series.Count, 2);
                }

                if (xRangeOkCount != chart.Series.Count)
                {
                    result.Errors.Add($"XSeries chưa đúng sau khi Switch Row/Column ({xRangeOkCount}/{chart.Series.Count}).");
                }
                if (yRangeOkCount != chart.Series.Count || headerOkCount != chart.Series.Count)
                {
                    result.Errors.Add($"YSeries/Header chưa đúng sau khi Switch Row/Column (Y={yRangeOkCount}/{chart.Series.Count}, Header={headerOkCount}/{chart.Series.Count}).");
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
