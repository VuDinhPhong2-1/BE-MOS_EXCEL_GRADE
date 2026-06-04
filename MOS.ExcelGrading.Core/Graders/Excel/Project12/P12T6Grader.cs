using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.Project12
{
    public class P12T6Grader : ITaskGrader
    {
        public string TaskId => "P12-T6";
        public string TaskName => "Inventory chart: hien thi tieu de + Data Labels vi tri outEnd";
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
                var chart = studentSheet.Workbook.Worksheets
                    .SelectMany(ws => ws.Drawings.OfType<ExcelChart>())
                    .FirstOrDefault(c =>
                    {
                        var firstSeries = c.Series.FirstOrDefault();
                        return firstSeries != null
                               && P12GraderHelpers.IsRangeMatch(firstSeries.XSeries?.ToString(), "Prices!B5:B25")
                               && P12GraderHelpers.IsRangeMatch(firstSeries.Series?.ToString(), "Prices!G5:G25");
                    });

                if (chart == null)
                {
                    result.Errors.Add("Không tìm thấy chart Inventory Level % can kiem tra.");
                    return result;
                }

                score += 1m;
                result.Details.Add("Tìm thấy dung chart Inventory Level %.");

                var xml = chart.ChartXml;
                var ns = new XmlNamespaceManager(xml.NameTable);
                ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

                var titleNode = xml.SelectSingleNode("//c:chart/c:title", ns);
                if (titleNode != null)
                {
                    score += 1m;
                    result.Details.Add("Chart da hien thi tieu de.");
                }
                else
                {
                    result.Errors.Add("Chart chua hien thi tieu de.");
                }

                var dLblPos = xml.SelectSingleNode("//c:ser/c:dLbls/c:dLblPos", ns)?.Attributes?["val"]?.Value ?? string.Empty;
                if (string.Equals(dLblPos, "outEnd", StringComparison.OrdinalIgnoreCase))
                {
                    score += 1m;
                    result.Details.Add("Data labels dung vi tri bên phải cột (outEnd).");
                }
                else
                {
                    result.Errors.Add($"Vi tri data label chưa đúng. Hiện tại: '{dLblPos}'.");
                }

                var showVal = xml.SelectSingleNode("//c:ser/c:dLbls/c:showVal", ns)?.Attributes?["val"]?.Value ?? string.Empty;
                if (string.Equals(showVal, "1", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Data labels da hien thi giá trị.");
                }
                else
                {
                    result.Errors.Add("Data labels chua hien thi giá trị.");
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



