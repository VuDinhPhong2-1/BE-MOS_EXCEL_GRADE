using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Xml;

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
                // Tìm sheet "Farmers & Market"
                var targetSheet = studentSheet.Workbook.Worksheets["Farmers & Market"]
                               ?? studentSheet.Workbook.Worksheets["Farmers & Markets"]
                               ?? studentSheet.Workbook.Worksheets["Farmer & Market"];

                if (targetSheet == null)
                {
                    result.Errors.Add("❌ Không tìm thấy sheet 'Farmers & Market'");
                    result.Score = 0;
                    return result;
                }

                decimal score = 0;
                bool foundChart = false;

                // Kiểm tra tất cả charts trong sheet
                foreach (var drawing in targetSheet.Drawings)
                {
                    if (drawing is ExcelChart chart)
                    {
                        Console.WriteLine($"=== Found Chart: {chart.Name} ===");
                        Console.WriteLine($"Chart Type: {chart.ChartType}");

                        // Rule 1: Kiểm tra loại chart (3 điểm)
                        bool isPie3D = CheckIfPie3D(chart);
                        if (isPie3D)
                        {
                            foundChart = true;
                            score += 3m;
                            result.Details.Add($"✓ Biểu đồ Pie 3D: {chart.Name}");

                            // Rule 2: Kiểm tra vị trí (2 điểm)
                            var position = CheckChartPosition(chart);
                            if (position.IsInRange)
                            {
                                score += 2m;
                                result.Details.Add($"✓ Vị trí đúng: {position.TopLeft} đến {position.BottomRight}");
                            }
                            else
                            {
                                result.Errors.Add($"❌ Vị trí sai: {position.TopLeft} (cần J2-P15)");
                            }

                            // Rule 3: Kiểm tra dữ liệu (3 điểm)
                            var dataCheck = CheckChartData(chart, targetSheet);
                            if (dataCheck.HasCorrectData)
                            {
                                score += 3m;
                                result.Details.Add($"✓ Dữ liệu đúng: Product và Total");
                                if (!string.IsNullOrEmpty(dataCheck.DataRange))
                                {
                                    result.Details.Add($"  • Vùng dữ liệu: {dataCheck.DataRange}");
                                }
                            }
                            else
                            {
                                score += dataCheck.PartialScore;
                                if (dataCheck.PartialScore > 0)
                                {
                                    result.Details.Add($"⚠ Dữ liệu một phần đúng ({dataCheck.PartialScore} điểm)");
                                }
                                else
                                {
                                    result.Errors.Add("❌ Dữ liệu không đúng (cần cột Product và Total)");
                                }
                            }

                            break; // Chỉ chấm chart đầu tiên tìm thấy
                        }
                    }
                }

                if (!foundChart)
                {
                    result.Errors.Add("❌ Không tìm thấy biểu đồ Pie 3D");
                }

                result.Score = score;
            }
            catch (Exception ex)
            {
                result.Errors.Add($"❌ Lỗi: {ex.Message}");
                result.Score = 0;
            }

            return result;
        }

        private bool CheckIfPie3D(ExcelChart chart)
        {
            try
            {
                // Kiểm tra các loại Pie 3D
                var validPie3DTypes = new[]
                {
                    eChartType.Pie3D,
                    eChartType.PieExploded3D
                };

                if (validPie3DTypes.Contains(chart.ChartType))
                {
                    Console.WriteLine($"✅ Chart type is 3D Pie: {chart.ChartType}");
                    return true;
                }

                // Kiểm tra qua XML nếu cần
                var chartXml = chart.ChartXml;
                var nsManager = new XmlNamespaceManager(new NameTable());
                nsManager.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

                var pie3DNode = chartXml.SelectSingleNode("//c:pie3DChart", nsManager);
                if (pie3DNode != null)
                {
                    Console.WriteLine("✅ Found pie3DChart in XML");
                    return true;
                }

                Console.WriteLine($"❌ Not a 3D Pie chart: {chart.ChartType}");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error checking chart type: {ex.Message}");
                return false;
            }
        }

        private (bool IsInRange, string TopLeft, string BottomRight) CheckChartPosition(ExcelChart chart)
        {
            try
            {
                // Lấy vị trí của chart
                int fromRow = chart.From.Row + 1; // EPPlus uses 0-based
                int fromCol = chart.From.Column + 1;
                int toRow = chart.To.Row + 1;
                int toCol = chart.To.Column + 1;

                string topLeft = GetCellAddress(fromRow, fromCol);
                string bottomRight = GetCellAddress(toRow, toCol);

                Console.WriteLine($"Chart position: {topLeft} to {bottomRight}");

                // Kiểm tra vùng J2:P15
                // J = cột 10, P = cột 16
                // Cho phép sai lệch ±1 ô
                bool isInRange = fromCol >= 9 && fromCol <= 11 &&  // J ± 1
                                fromRow >= 1 && fromRow <= 3 &&     // Row 2 ± 1
                                toCol >= 15 && toCol <= 17 &&       // P ± 1  
                                toRow >= 14 && toRow <= 16;         // Row 15 ± 1

                return (isInRange, topLeft, bottomRight);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error checking position: {ex.Message}");
                return (false, "Unknown", "Unknown");
            }
        }

        private (bool HasCorrectData, decimal PartialScore, string DataRange) CheckChartData(ExcelChart chart, ExcelWorksheet sheet)
        {
            try
            {
                bool hasProductColumn = false;
                bool hasTotalColumn = false;
                string dataRange = "";

                foreach (var series in chart.Series)
                {
                    if (series is ExcelPieChartSerie pieSeries)
                    {
                        // ✅ SỬA: Dùng HeaderAddress trực tiếp (không có .Address)
                        var headerAddress = pieSeries.HeaderAddress?.Address;
                        if (!string.IsNullOrEmpty(headerAddress))
                        {
                            Console.WriteLine($"Header: {headerAddress}");
                            var headerValue = sheet.Cells[headerAddress].Text;
                            if (headerValue.Contains("Product", StringComparison.OrdinalIgnoreCase))
                            {
                                hasProductColumn = true;
                            }
                        }

                        // Categories (Product)
                        var xLabels = pieSeries.XSeries?.ToString();
                        if (!string.IsNullOrEmpty(xLabels))
                        {
                            Console.WriteLine($"Categories: {xLabels}");
                            if (xLabels.Contains("$B$", StringComparison.OrdinalIgnoreCase) ||
                                xLabels.Contains("$B:", StringComparison.OrdinalIgnoreCase) ||
                                CheckIfProductColumn(sheet, xLabels))
                            {
                                hasProductColumn = true;
                            }
                        }

                        // Values (Total)
                        var values = pieSeries.Series?.ToString();
                        if (!string.IsNullOrEmpty(values))
                        {
                            Console.WriteLine($"Values: {values}");
                            dataRange = values;

                            if (values.Contains("$F$", StringComparison.OrdinalIgnoreCase) ||
                                values.Contains("$F:", StringComparison.OrdinalIgnoreCase) ||
                                values.Contains("Total", StringComparison.OrdinalIgnoreCase) ||
                                CheckIfTotalColumn(sheet, values))
                            {
                                hasTotalColumn = true;
                            }
                        }
                    }
                } // ✅ THÊM: Dấu } đóng foreach

                // Fallback to XML check
                if (!hasProductColumn || !hasTotalColumn)
                {
                    var xmlCheck = CheckChartDataViaXML(chart);
                    hasProductColumn = hasProductColumn || xmlCheck.HasProduct;
                    hasTotalColumn = hasTotalColumn || xmlCheck.HasTotal;
                    if (string.IsNullOrEmpty(dataRange))
                    {
                        dataRange = xmlCheck.Range;
                    }
                }

                Console.WriteLine($"Data check - Product: {hasProductColumn}, Total: {hasTotalColumn}");

                decimal partialScore = 0;
                if (hasProductColumn) partialScore += 1.5m;
                if (hasTotalColumn) partialScore += 1.5m;

                bool hasCorrectData = hasProductColumn && hasTotalColumn;
                return (hasCorrectData, partialScore, dataRange);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error checking data: {ex.Message}");
                return (false, 0, "");
            }
        }

        private bool CheckIfProductColumn(ExcelWorksheet sheet, string range)
        {
            try
            {
                // Kiểm tra header của cột
                var cells = sheet.Cells[range];
                if (cells.Start.Row > 1)
                {
                    var headerCell = sheet.Cells[1, cells.Start.Column];
                    return headerCell.Text.Contains("Product", StringComparison.OrdinalIgnoreCase);
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        private bool CheckIfTotalColumn(ExcelWorksheet sheet, string range)
        {
            try
            {
                // Kiểm tra header của cột
                var cells = sheet.Cells[range];
                if (cells.Start.Row > 1)
                {
                    var headerCell = sheet.Cells[1, cells.Start.Column];
                    return headerCell.Text.Contains("Total", StringComparison.OrdinalIgnoreCase);
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        private (bool HasProduct, bool HasTotal, string Range) CheckChartDataViaXML(ExcelChart chart)
        {
            try
            {
                var chartXml = chart.ChartXml;
                var nsManager = new XmlNamespaceManager(new NameTable());
                nsManager.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

                bool hasProduct = false;
                bool hasTotal = false;
                string range = "";

                // Tìm cat (categories) và val (values)
                var catNode = chartXml.SelectSingleNode("//c:cat/c:strRef/c:f", nsManager);
                var valNode = chartXml.SelectSingleNode("//c:val/c:numRef/c:f", nsManager);

                if (catNode != null)
                {
                    var catRange = catNode.InnerText;
                    Console.WriteLine($"XML Categories: {catRange}");
                    if (catRange.Contains("$A$", StringComparison.OrdinalIgnoreCase))
                    {
                        hasProduct = true;
                    }
                }

                if (valNode != null)
                {
                    var valRange = valNode.InnerText;
                    Console.WriteLine($"XML Values: {valRange}");
                    range = valRange;
                    if (valRange.Contains("$B$", StringComparison.OrdinalIgnoreCase))
                    {
                        hasTotal = true;
                    }
                }

                return (hasProduct, hasTotal, range);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error checking XML: {ex.Message}");
                return (false, false, "");
            }
        }

        private string GetCellAddress(int row, int col)
        {
            string columnLetter = "";
            while (col > 0)
            {
                col--;
                columnLetter = (char)('A' + col % 26) + columnLetter;
                col /= 26;
            }
            return $"{columnLetter}{row}";
        }
    }
}
