using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System.Xml;

namespace MOS.ExcelGrading.Core.Graders.Project07
{
    public class P07T3Grader : ITaskGrader
    {
        public string TaskId => "P07-T3";
        public string TaskName => "Tea chart Layout 9 + Vertical Axis Title='Price' + remove Horizontal Axis Title";
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
                var ws = P07GraderHelpers.GetSheet(studentSheet, "Tea");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Tea'.");
                    return result;
                }

                var chart = ws.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Khong tim thay chart tren sheet Tea.");
                    return result;
                }

                decimal score = 1m; // Tim thay chart.

                var yTitle = chart.YAxis?.Title?.Text ?? string.Empty;
                if (string.Equals(yTitle, "Price", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Primary Vertical Axis Title da dat dung 'Price'.");
                }
                else
                {
                    result.Errors.Add($"Primary Vertical Axis Title chua dung. Hien tai: '{yTitle}'.");
                }

                var xTitle = chart.XAxis?.Title?.Text ?? string.Empty;
                var xTitleTechnicalClean = xTitle
                    .Replace("\r", string.Empty, StringComparison.Ordinal)
                    .Replace("\n", string.Empty, StringComparison.Ordinal)
                    .Replace("\t", string.Empty, StringComparison.Ordinal);
                if (xTitleTechnicalClean.Length == 0)
                {
                    score += 1m;
                    result.Details.Add("Horizontal Axis Title da duoc xoa.");
                }
                else
                {
                    result.Errors.Add($"Horizontal Axis Title chua duoc xoa. Hien tai: '{xTitle}'.");
                }

                var xml = chart.ChartXml;
                var ns = new XmlNamespaceManager(xml.NameTable);
                ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
                var legendPos = xml.SelectSingleNode("//c:legend/c:legendPos", ns)?.Attributes?["val"]?.Value;
                if (string.Equals(legendPos, "r", StringComparison.OrdinalIgnoreCase))
                {
                    score += 1m;
                    result.Details.Add("Chart legend o ben phai (dac trung Layout 9).");
                }
                else
                {
                    result.Errors.Add("Legend chart chua o ben phai, bo cuc chua khop Layout 9.");
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
