using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Xml;

namespace MOS.ExcelGrading.Core.Graders.Project09
{
    public class P09T5Grader : ITaskGrader
    {
        public string TaskId => "P09-T5";
        public string TaskName => "Subtotal by shirt color, page breaks, Grand Total";
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
                var dataSheet = studentSheet.Workbook.Worksheets["Shirt Orders"] ?? studentSheet;
                decimal score = 0;

                // Rule 1: Có Subtotal (3 điểm)
                bool hasSubtotal = CheckOutlineExists(dataSheet);
                if (hasSubtotal)
                {
                    score += 3m;
                    result.Details.Add("✓ Đã tạo Subtotal/Grouping");
                }
                else
                {
                    result.Errors.Add("❌ Chưa tạo Subtotal");
                }

                // Rule 2: Có Page Breaks (2 điểm)
                int pageBreakCount = CheckPageBreaks(dataSheet);

                if (pageBreakCount > 0)
                {
                    score += 2m;
                    result.Details.Add($"✓ Có {pageBreakCount} page breaks");
                }
                else
                {
                    result.Errors.Add("❌ Chưa có page breaks");
                }

                // Rule 3: Grand Total (3 điểm)
                var grandTotalResult = FindAndCheckGrandTotal(dataSheet);

                if (grandTotalResult.HasValue)
                {
                    if (grandTotalResult.HasFormula)
                    {
                        score += 3m;
                        result.Details.Add($"✓ Grand Total tại {grandTotalResult.CellAddress}: {grandTotalResult.Value:N0}");
                    }
                    else
                    {
                        score += 1.5m;
                        result.Details.Add($"⚠ Có giá trị tại {grandTotalResult.CellAddress} ({grandTotalResult.Value:N0}) nhưng không có công thức");
                    }
                }
                else
                {
                    result.Errors.Add("❌ Không tìm thấy Grand Total");
                }

                result.Score = score;
            }
            catch (Exception ex)
            {
                result.Errors.Add($"❌ Lỗi: {ex.Message}");
            }

            return result;
        }

        private bool CheckOutlineExists(ExcelWorksheet sheet)
        {
            try
            {
                if (sheet.Dimension == null) return false;

                Console.WriteLine("=== Checking Subtotal ===");
                int matchCount = 0;

                // Cách 1: OutlineLevel
                for (int row = 1; row <= sheet.Dimension.End.Row; row++)
                {
                    if (sheet.Row(row).OutlineLevel > 0)
                    {
                        matchCount++;
                        Console.WriteLine($"✅ [OutlineLevel] Found at row {row}");
                        break;
                    }
                }

                // Cách 2: Text "Total"
                int subtotalText = 0;
                for (int row = 1; row <= sheet.Dimension.End.Row; row++)
                {
                    for (int col = 1; col <= Math.Min(10, sheet.Dimension.End.Column); col++)
                    {
                        var cellValue = sheet.Cells[row, col].Text ?? "";
                        if ((cellValue.Contains("Total", StringComparison.OrdinalIgnoreCase) ||
                             cellValue.Contains("Subtotal", StringComparison.OrdinalIgnoreCase)) &&
                            !cellValue.Contains("Grand", StringComparison.OrdinalIgnoreCase))
                        {
                            subtotalText++;
                            Console.WriteLine($"✅ [Text] Found at row {row}, col {col}: {cellValue}");
                        }
                    }
                }
                if (subtotalText > 1) matchCount++;

                // Cách 3: Formula SUBTOTAL
                int subtotalFormula = 0;
                for (int row = 1; row <= sheet.Dimension.End.Row; row++)
                {
                    for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                    {
                        var formula = sheet.Cells[row, col].Formula;
                        if (!string.IsNullOrEmpty(formula) &&
                            formula.Contains("SUBTOTAL", StringComparison.OrdinalIgnoreCase))
                        {
                            subtotalFormula++;
                            Console.WriteLine($"✅ [Formula] Found at {sheet.Cells[row, col].Address}: {formula}");
                        }
                    }
                }
                if (subtotalFormula > 1) matchCount++;

                // Cách 4: XML Outline
                var worksheetXml = sheet.WorksheetXml;
                var nsManager = new XmlNamespaceManager(new NameTable());
                nsManager.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var sheetFormatNode = worksheetXml.SelectSingleNode("//x:sheetFormatPr", nsManager);
                if (sheetFormatNode != null &&
                    int.TryParse(sheetFormatNode.Attributes?["outlineLevelRow"]?.Value, out int olv) &&
                    olv > 0)
                {
                    matchCount++;
                    Console.WriteLine($"✅ [XML] Outline level: {olv}");
                }

                Console.WriteLine($"=== Subtotal Check: {matchCount}/4 ===");
                return matchCount >= 2;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error: {ex.Message}");
                return false;
            }
        }

        private int CheckPageBreaks(ExcelWorksheet sheet)
        {
            try
            {
                Console.WriteLine("=== Checking Page Breaks ===");

                var worksheetXml = sheet.WorksheetXml;
                var nsManager = new XmlNamespaceManager(new NameTable());
                nsManager.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                // Kiểm tra rowBreaks
                XmlNode? rowBreaksNode = worksheetXml.SelectSingleNode("//x:rowBreaks", nsManager)
                                      ?? worksheetXml.SelectSingleNode("//rowBreaks")
                                      ?? worksheetXml.SelectSingleNode("//*[local-name()='rowBreaks']");

                if (rowBreaksNode != null)
                {
                    Console.WriteLine("✅ Found rowBreaks node");

                    // Kiểm tra attribute manualBreakCount (quan trọng!)
                    var manualBreakCount = rowBreaksNode.Attributes?["manualBreakCount"]?.Value;
                    if (!string.IsNullOrEmpty(manualBreakCount) && int.TryParse(manualBreakCount, out int manualCount) && manualCount > 0)
                    {
                        Console.WriteLine($"✅ Found {manualCount} manual page breaks");
                        return manualCount;
                    }

                    // Kiểm tra attribute count
                    var countAttr = rowBreaksNode.Attributes?["count"]?.Value;
                    if (!string.IsNullOrEmpty(countAttr) && int.TryParse(countAttr, out int count) && count > 0)
                    {
                        Console.WriteLine($"✅ Found {count} page breaks via count");
                        return count;
                    }

                    // Đếm các node brk
                    var breakNodes = rowBreaksNode.SelectNodes("x:brk", nsManager)
                                  ?? rowBreaksNode.SelectNodes("brk")
                                  ?? rowBreaksNode.SelectNodes("*[local-name()='brk']");

                    if (breakNodes != null && breakNodes.Count > 0)
                    {
                        Console.WriteLine($"✅ Found {breakNodes.Count} break nodes");
                        return breakNodes.Count;
                    }

                    // QUAN TRỌNG: Nếu có rowBreaks node = đã setup page breaks (dù rỗng)
                    // Trong Excel, khi bật "Page Break Between Groups", nó tạo rowBreaks rỗng
                    // và breaks thực tế được tính tự động dựa trên grouping
                    if (CheckOutlineExists(sheet))
                    {
                        Console.WriteLine("✅ Found rowBreaks + Outline → Auto page breaks between groups");

                        // Đếm số lượng group (subtotal rows)
                        int groupCount = CountSubtotalGroups(sheet);
                        if (groupCount > 0)
                        {
                            Console.WriteLine($"✅ Estimated {groupCount} page breaks from {groupCount} groups");
                            return groupCount;
                        }
                    }

                    Console.WriteLine("⚠ Found rowBreaks node but empty → assuming 1 page break setup");
                    return 1; // Có node = có setup, tối thiểu 1 break
                }

                Console.WriteLine("❌ No rowBreaks node found");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error: {ex.Message}");
                return 0;
            }
        }

        // Hàm helper đếm số group
        private int CountSubtotalGroups(ExcelWorksheet sheet)
        {
            int count = 0;
            if (sheet.Dimension == null) return 0;

            for (int row = 1; row <= sheet.Dimension.End.Row; row++)
            {
                for (int col = 1; col <= Math.Min(10, sheet.Dimension.End.Column); col++)
                {
                    var cellValue = sheet.Cells[row, col].Text ?? "";
                    if (cellValue.Contains("Total", StringComparison.OrdinalIgnoreCase) &&
                        !cellValue.Contains("Grand", StringComparison.OrdinalIgnoreCase))
                    {
                        count++;
                        break; // Chỉ đếm 1 lần per row
                    }
                }
            }
            return count;
        }

        private GrandTotalResult FindAndCheckGrandTotal(ExcelWorksheet sheet)
        {
            var result = new GrandTotalResult();
            try
            {
                if (sheet.Dimension == null) return result;

                Console.WriteLine("=== Finding Grand Total ===");

                for (int row = sheet.Dimension.End.Row; row >= 1; row--)
                {
                    for (int col = 1; col <= Math.Min(sheet.Dimension.End.Column, 10); col++)
                    {
                        var cellValue = sheet.Cells[row, col].Text ?? "";
                        if (cellValue.Contains("Grand", StringComparison.OrdinalIgnoreCase) &&
                            cellValue.Contains("Total", StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine($"✅ Found 'Grand Total' text at row {row}, col {col}");

                            for (int c = sheet.Dimension.End.Column; c >= 1; c--)
                            {
                                var val = sheet.Cells[row, c].Value;
                                if (val != null && double.TryParse(val.ToString(), out double num))
                                {
                                    result.HasValue = true;
                                    result.Value = num;
                                    result.CellAddress = sheet.Cells[row, c].Address;

                                    var formula = sheet.Cells[row, c].Formula;
                                    if (!string.IsNullOrEmpty(formula))
                                    {
                                        result.HasFormula = true;
                                        Console.WriteLine($"✅ Grand Total at {result.CellAddress}: {num:N0} (formula: {formula})");
                                    }
                                    else
                                    {
                                        Console.WriteLine($"⚠ Grand Total at {result.CellAddress}: {num:N0} (no formula)");
                                    }

                                    return result;
                                }
                            }
                        }
                    }
                }

                Console.WriteLine("❌ Grand Total not found");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error: {ex.Message}");
            }
            return result;
        }

        private class GrandTotalResult
        {
            public bool HasValue { get; set; }
            public double Value { get; set; }
            public bool HasFormula { get; set; }
            public string CellAddress { get; set; } = string.Empty;
        }
    }
}
