using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Xml;

namespace MOS.ExcelGrading.Core.Graders.Project02
{
    public class P02T6Grader : ITaskGrader
    {
        public string TaskId => "P02-T6";
        public string TaskName => "Doi chart layout thanh Layout 3 tren New Policy";
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
                var ws = studentSheet.Workbook.Worksheets["New Policy"];
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'New Policy'");
                    return result;
                }

                var chart = ws.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Khong tim thay bieu do tren sheet New Policy");
                    return result;
                }

                decimal score = 0;
                score += 1m;

                var xml = chart.ChartXml;
                var ns = new XmlNamespaceManager(xml.NameTable);
                ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

                var legendPos = xml.SelectSingleNode("//c:legend/c:legendPos", ns)?.Attributes?["val"]?.Value;
                if (string.Equals(legendPos, "b", StringComparison.OrdinalIgnoreCase))
                {
                    score += 1m;
                    result.Details.Add("Legend dang o vi tri Bottom");
                }
                else
                {
                    result.Errors.Add("Legend chua o vi tri Bottom (Layout 3)");
                }

                var titleNode = xml.SelectSingleNode("//c:chart/c:title", ns);
                var autoTitleDeleted = xml.SelectSingleNode("//c:chart/c:autoTitleDeleted", ns)?.Attributes?["val"]?.Value;
                if (titleNode != null || autoTitleDeleted == "0")
                {
                    score += 1m;
                    result.Details.Add("Chart title duoc hien thi");
                }
                else
                {
                    result.Errors.Add("Chart title chua duoc hien thi nhu Layout 3");
                }

                var hasDataTable = xml.SelectSingleNode("//c:plotArea/c:dTable", ns) != null;
                var hasManualLayout = xml.SelectSingleNode("//c:plotArea/c:layout/c:manualLayout", ns) != null;
                if (!hasDataTable && !hasManualLayout)
                {
                    score += 1m;
                    result.Details.Add("Bo cuc plot area phu hop voi Layout 3");
                }
                else
                {
                    result.Errors.Add("Plot area van con cau hinh cu (data table/manual layout)");
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

