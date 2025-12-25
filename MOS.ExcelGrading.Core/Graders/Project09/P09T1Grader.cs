using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Xml;

namespace MOS.ExcelGrading.Core.Graders.Project09
{
    public class P09T1Grader : ITaskGrader
    {
        public string TaskId => "P09-T1";
        public string TaskName => "Apply pattern fill to chart (10% plot area, 50% chart area)";
        public decimal MaxScore => 5;

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
                var summarySheet = studentSheet.Workbook.Worksheets["Summary"];

                if (summarySheet == null)
                {
                    result.Errors.Add("❌ Không tìm thấy sheet 'Summary'");
                    return result;
                }

                var summaryChart = summarySheet.Drawings.FirstOrDefault() as ExcelChart;

                if (summaryChart == null)
                {
                    result.Errors.Add("❌ Không tìm thấy biểu đồ");
                    return result;
                }

                decimal score = 0;

                // Rule 1: Plot Area có pattern fill 10% (2.5 điểm)
                var (hasPlotFill, plotPattern) = CheckPlotAreaFill(summaryChart);

                if (hasPlotFill)
                {
                    if (plotPattern == "pct10")
                    {
                        score += 2.5m;
                        result.Details.Add($"✓ Plot area có pattern fill 10%");
                    }
                    else
                    {
                        score += 1m;
                        result.Errors.Add($"⚠ Plot area có pattern fill nhưng type là '{plotPattern}' (mong đợi: pct10)");
                    }
                }
                else
                {
                    result.Errors.Add("❌ Plot area chưa có pattern fill");
                }

                // Rule 2: Chart Area có pattern fill 50% (2.5 điểm)
                var (hasChartFill, chartPattern) = CheckChartAreaFill(summaryChart);

                if (hasChartFill)
                {
                    if (chartPattern == "pct50")
                    {
                        score += 2.5m;
                        result.Details.Add($"✓ Chart area có pattern fill 50%");
                    }
                    else
                    {
                        score += 1m;
                        result.Errors.Add($"⚠ Chart area có pattern fill nhưng type là '{chartPattern}' (mong đợi: pct50)");
                    }
                }
                else
                {
                    result.Errors.Add("❌ Chart area chưa có pattern fill");
                }

                result.Score = score;
            }
            catch (Exception ex)
            {
                result.Errors.Add($"❌ Lỗi: {ex.Message}");
            }

            return result;
        }

        private (bool HasFill, string PatternType) CheckPlotAreaFill(ExcelChart chart)
        {
            try
            {
                var xml = chart.ChartXml;
                var nsManager = CreateNamespaceManager();

                var plotAreaNode = xml.SelectSingleNode("//c:plotArea", nsManager);
                if (plotAreaNode == null) return (false, null);

                var spPrNode = plotAreaNode.SelectSingleNode("c:spPr", nsManager);
                if (spPrNode == null) return (false, null);

                var pattFillNode = spPrNode.SelectSingleNode("a:pattFill", nsManager);

                if (pattFillNode != null)
                {
                    var prstAttr = pattFillNode.Attributes?["prst"];

                    if (prstAttr != null)
                    {
                        var patternType = prstAttr.Value;
                        Console.WriteLine($"✅ Plot area pattern type: {patternType}");
                        return (true, patternType);
                    }
                }

                return (false, null);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error: {ex.Message}");
                return (false, null);
            }
        }

        private (bool HasFill, string PatternType) CheckChartAreaFill(ExcelChart chart)
        {
            try
            {
                var xml = chart.ChartXml;
                var nsManager = CreateNamespaceManager();

                var chartSpaceNode = xml.SelectSingleNode("//c:chartSpace", nsManager);
                if (chartSpaceNode == null) return (false, null);

                var spPrNode = chartSpaceNode.SelectSingleNode("c:spPr", nsManager);
                if (spPrNode == null) return (false, null);

                var pattFillNode = spPrNode.SelectSingleNode("a:pattFill", nsManager);

                if (pattFillNode != null)
                {
                    var prstAttr = pattFillNode.Attributes?["prst"];

                    if (prstAttr != null)
                    {
                        var patternType = prstAttr.Value;
                        Console.WriteLine($"✅ Chart area pattern type: {patternType}");
                        return (true, patternType);
                    }
                }

                return (false, null);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error: {ex.Message}");
                return (false, null);
            }
        }

        private XmlNamespaceManager CreateNamespaceManager()
        {
            var nsManager = new XmlNamespaceManager(new NameTable());
            nsManager.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            nsManager.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            return nsManager;
        }
    }
}
