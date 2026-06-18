using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Xml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project09
{
    public class P09T3Grader : ITaskGrader
    {
        public string TaskId => "P09-T3";
        public string TaskName => "Hien thi Legend ben phai, cho phep tran";
        public decimal MaxScore => 3;

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
                // L?y sheet Summary
                var summarySheet = studentSheet.Workbook.Worksheets["Summary"];
                
                if (summarySheet == null)
                {
                    result.Errors.Add("? Không tìm th?y sheet 'Summary'");
                    return result;
                }

                // L?y chart d?u tiên
                var summaryChart = summarySheet.Drawings.FirstOrDefault() as ExcelChart;

                if (summaryChart == null)
                {
                    result.Errors.Add("? Không tìm th?y bi?u d? trong sheet Summary");
                    return result;
                }

                decimal score = 0;

                // Rule 1: Legend position = Right (2 di?m)
                if (CheckLegendPosition(summaryChart))
                {
                    score += 2m;
                    result.Details.Add("? Legend ? bên ph?i");
                }
                else
                {
                    result.Errors.Add("? Legend không ? bên ph?i");
                }

                // Rule 2: Legend overlay = true (1 di?m)
                if (CheckLegendOverlay(summaryChart))
                {
                    score += 1m;
                    result.Details.Add("? Legend cho phép tràn qua bi?u d?");
                }
                else
                {
                    result.Errors.Add("? Legend chua cho phép tràn qua bi?u d?");
                }

                result.Score = score;
            }
            catch (Exception ex)
            {
                result.Errors.Add( $"? L?i: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Ki?m tra legend position = Right
        /// </summary>
        private bool CheckLegendPosition(ExcelChart chart)
        {
            try
            {
                // Ki?m tra qua EPPlus API
                if (chart.Legend != null && chart.Legend.Position == eLegendPosition.Right)
                {
                    Console.WriteLine("? Legend position: Right");
                    return true;
                }

                // N?u API không có, ki?m tra qua XML
                var xml = chart.ChartXml;
                var nsManager = CreateNamespaceManager();

                var legendNode = xml.SelectSingleNode("//c:legend", nsManager);
                
                if (legendNode == null)
                {
                    Console.WriteLine("? Legend node not found");
                    return false;
                }

                var legendPosNode = legendNode.SelectSingleNode("c:legendPos", nsManager);
                
                if (legendPosNode != null)
                {
                    var valAttr = legendPosNode.Attributes?["val"];
                    
                    if (valAttr != null)
                    {
                        var position = valAttr.Value;
                        Console.WriteLine( $"Legend position: {position}");
                        
                        // "r" = right
                        return position == "r";
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine( $"? Error checking legend position: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Ki?m tra legend overlay = true
        /// </summary>
        private bool CheckLegendOverlay(ExcelChart chart)
        {
            try
            {
                var xml = chart.ChartXml;
                var nsManager = CreateNamespaceManager();

                var legendNode = xml.SelectSingleNode("//c:legend", nsManager);
                
                if (legendNode == null)
                {
                    Console.WriteLine("? Legend node not found");
                    return false;
                }

                // Tìm node overlay
                var overlayNode = legendNode.SelectSingleNode("c:overlay", nsManager);
                
                if (overlayNode != null)
                {
                    var valAttr = overlayNode.Attributes?["val"];
                    
                    if (valAttr != null)
                    {
                        var overlayValue = valAttr.Value;
                        Console.WriteLine( $"Legend overlay: {overlayValue}");
                        
                        // "1" ho?c "true" = overlay enabled
                        return overlayValue == "1" || overlayValue.ToLower() == "true";
                    }
                    else
                    {
                        // N?u có node overlay nhung không có val attribute
                        // ? M?c d?nh là true
                        Console.WriteLine("? Legend overlay: true (default)");
                        return true;
                    }
                }
                else
                {
                    Console.WriteLine("? Overlay node not found");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine( $"? Error checking legend overlay: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// T?o XmlNamespaceManager
        /// </summary>
        private XmlNamespaceManager CreateNamespaceManager()
        {
            var nsManager = new XmlNamespaceManager(new NameTable());
            nsManager.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            nsManager.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            return nsManager;
        }
    }
}

// minor-sync: non-functional graders update




