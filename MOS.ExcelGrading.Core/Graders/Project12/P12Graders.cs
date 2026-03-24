using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.Project12
{
    internal static class P12GraderHelpers
    {
        public static ExcelWorksheet? GetSheet(ExcelWorkbook workbook, string name)
        {
            return workbook.Worksheets.FirstOrDefault(sheet =>
                string.Equals((sheet.Name ?? string.Empty).Trim(), name.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        public static string NormalizeRange(string? range)
        {
            var text = (range ?? string.Empty)
                .Trim()
                .Replace("=", string.Empty, StringComparison.Ordinal)
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Replace("'", string.Empty, StringComparison.Ordinal)
                .Replace(" ", string.Empty, StringComparison.Ordinal);

            var excl = text.LastIndexOf('!');
            if (excl >= 0 && excl + 1 < text.Length)
            {
                text = text[(excl + 1)..];
            }

            return text.ToUpperInvariant();
        }

        public static bool IsRangeMatch(string? actual, string expected)
        {
            return string.Equals(NormalizeRange(actual), NormalizeRange(expected), StringComparison.OrdinalIgnoreCase);
        }

        public static string NormalizeFormula(string? formula)
        {
            return (formula ?? string.Empty)
                .Trim()
                .Replace("=", string.Empty, StringComparison.Ordinal)
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Replace(" ", string.Empty, StringComparison.Ordinal)
                .Replace("_xlfn.", string.Empty, StringComparison.OrdinalIgnoreCase)
                .ToUpperInvariant();
        }

        public static bool HasMergeRange(ExcelWorksheet worksheet, string expectedRange)
        {
            return worksheet.MergedCells.Any(range => IsRangeMatch(range, expectedRange));
        }

        public static ExcelTable? FindTableByAddress(ExcelWorksheet worksheet, string expectedAddress)
        {
            return worksheet.Tables.FirstOrDefault(table => IsRangeMatch(table.Address.Address, expectedAddress));
        }

        public static string GetSingleFilterValue(ExcelTable table, int colId)
        {
            var tableXml = table.TableXml;
            var ns = new XmlNamespaceManager(tableXml.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            var filterNode = tableXml.SelectSingleNode(
                $"//x:autoFilter/x:filterColumn[@colId='{colId}']/x:filters/x:filter",
                ns);
            return filterNode?.Attributes?["val"]?.Value?.Trim() ?? string.Empty;
        }
    }

    public class P12T1Grader : ITaskGrader
    {
        public string TaskId => "P12-T1";
        public string TaskName => "Range: merge E7:F7";
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
                var ws = P12GraderHelpers.GetSheet(studentSheet.Workbook, "Range");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Range'.");
                    return result;
                }

                decimal score = 0m;
                if (P12GraderHelpers.HasMergeRange(ws, "E7:F7"))
                {
                    score += 2m;
                    result.Details.Add("Da merge dung vung E7:F7.");
                }
                else
                {
                    result.Errors.Add("Chua merge dung vung E7:F7.");
                }

                if (!P12GraderHelpers.HasMergeRange(ws, "E8:F8"))
                {
                    score += 1m;
                    result.Details.Add("Khong con merge cu E8:F8.");
                }
                else
                {
                    result.Errors.Add("Van con merge cu E8:F8.");
                }

                var styleMoved = ws.Cells["E7"].StyleID > 0 && ws.Cells["F7"].StyleID > 0;
                if (styleMoved)
                {
                    score += 1m;
                    result.Details.Add("Dinh dang cua o merge duoc giu lai.");
                }
                else
                {
                    result.Errors.Add("Dinh dang o merge E7:F7 khong hop le.");
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

    public class P12T2Grader : ITaskGrader
    {
        public string TaskId => "P12-T2";
        public string TaskName => "Prices: apply Title style to A1";
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
                var ws = P12GraderHelpers.GetSheet(studentSheet.Workbook, "Prices");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Prices'.");
                    return result;
                }

                var cell = ws.Cells["A1"];
                decimal score = 0m;

                if (string.Equals(cell.Style.Font.Name ?? string.Empty, "Century Gothic", StringComparison.OrdinalIgnoreCase))
                {
                    score += 1m;
                    result.Details.Add("Font A1 dung: Century Gothic.");
                }
                else
                {
                    result.Errors.Add($"Font A1 chua dung. Hien tai: '{cell.Style.Font.Name}'.");
                }

                if (Math.Abs(cell.Style.Font.Size - 18f) <= 0.1f)
                {
                    score += 1m;
                    result.Details.Add("Co chu A1 dung: 18.");
                }
                else
                {
                    result.Errors.Add($"Co chu A1 chua dung. Hien tai: {cell.Style.Font.Size:0.##}.");
                }

                if (cell.Style.HorizontalAlignment == ExcelHorizontalAlignment.Left)
                {
                    score += 1m;
                    result.Details.Add("Can le ngang A1 dung: Left.");
                }
                else
                {
                    result.Errors.Add($"Can le ngang A1 chua dung. Hien tai: {cell.Style.HorizontalAlignment}.");
                }

                if (cell.Style.VerticalAlignment == ExcelVerticalAlignment.Center)
                {
                    score += 1m;
                    result.Details.Add("Can le doc A1 dung: Center.");
                }
                else
                {
                    result.Errors.Add($"Can le doc A1 chua dung. Hien tai: {cell.Style.VerticalAlignment}.");
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

    public class P12T3Grader : ITaskGrader
    {
        public string TaskId => "P12-T3";
        public string TaskName => "Orders: filter by The House of Alpine Skiing";
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
                var ws = P12GraderHelpers.GetSheet(studentSheet.Workbook, "Orders");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Orders'.");
                    return result;
                }

                decimal score = 0m;
                var table = P12GraderHelpers.FindTableByAddress(ws, "A1:E412");
                if (table != null)
                {
                    score += 1m;
                    result.Details.Add("Tim thay table Orders dung range A1:E412.");
                }
                else
                {
                    result.Errors.Add("Khong tim thay table Orders range A1:E412.");
                    result.Score = score;
                    return result;
                }

                var filteredValue = P12GraderHelpers.GetSingleFilterValue(table, 0);
                if (string.Equals(filteredValue, "The House of Alpine Skiing", StringComparison.Ordinal))
                {
                    score += 2m;
                    result.Details.Add("Dieu kien filter cot dau dung gia tri yeu cau.");
                }
                else
                {
                    result.Errors.Add($"Gia tri filter chua dung. Hien tai: '{filteredValue}'.");
                }

                if (table.TableStyle == TableStyles.Light18)
                {
                    score += 1m;
                    result.Details.Add("Table style Orders dung: TableStyleLight18.");
                }
                else
                {
                    result.Errors.Add($"Table style Orders chua dung. Hien tai: {table.TableStyle}.");
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

    public class P12T4Grader : ITaskGrader
    {
        public string TaskId => "P12-T4";
        public string TaskName => "Prices: Tax formula uses Unit price * L$2";
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
                var ws = P12GraderHelpers.GetSheet(studentSheet.Workbook, "Prices");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Prices'.");
                    return result;
                }

                decimal score = 0m;
                var table = P12GraderHelpers.FindTableByAddress(ws, "A4:L25");
                if (table == null)
                {
                    result.Errors.Add("Khong tim thay table Prices A4:L25.");
                    return result;
                }

                var taxColumn = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name, "Tax", StringComparison.OrdinalIgnoreCase));
                if (taxColumn == null)
                {
                    result.Errors.Add("Khong tim thay cot 'Tax' trong table Prices.");
                    return result;
                }

                var normalizedFormula = P12GraderHelpers.NormalizeFormula(taxColumn.CalculatedColumnFormula);
                var expectedFormula = P12GraderHelpers.NormalizeFormula("Table2[[#This Row],[Unit price]]*L$2");
                if (string.Equals(normalizedFormula, expectedFormula, StringComparison.Ordinal))
                {
                    score += 3m;
                    result.Details.Add("Cong thuc cot Tax dung chinh xac theo structured reference.");
                }
                else
                {
                    result.Errors.Add($"Cong thuc cot Tax chua dung. Hien tai: '{taxColumn.CalculatedColumnFormula}'.");
                }

                if (normalizedFormula.Contains("TABLE2[[#THISROW],[UNITPRICE]]", StringComparison.Ordinal)
                    && normalizedFormula.Contains("L2", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Cong thuc Tax dung nguon cot Unit price va o L$2.");
                }
                else
                {
                    result.Errors.Add("Cong thuc Tax khong dung nguon du lieu Unit price va L$2.");
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

    public class P12T5Grader : ITaskGrader
    {
        public string TaskId => "P12-T5";
        public string TaskName => "Prices: Inventory Notice IF <15% then Low else blank";
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
                var ws = P12GraderHelpers.GetSheet(studentSheet.Workbook, "Prices");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Prices'.");
                    return result;
                }

                decimal score = 0m;
                var table = P12GraderHelpers.FindTableByAddress(ws, "A4:L25");
                if (table == null)
                {
                    result.Errors.Add("Khong tim thay table Prices A4:L25.");
                    return result;
                }

                var noticeColumn = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name, "Inventory Notice", StringComparison.OrdinalIgnoreCase));
                if (noticeColumn == null)
                {
                    result.Errors.Add("Khong tim thay cot 'Inventory Notice'.");
                    return result;
                }

                var normalizedFormula = P12GraderHelpers.NormalizeFormula(noticeColumn.CalculatedColumnFormula);
                var expectedFormula = P12GraderHelpers.NormalizeFormula("IF(Table2[[#This Row],[Inventory Level %]]<15%,\"Low\",\"\")");
                if (string.Equals(normalizedFormula, expectedFormula, StringComparison.Ordinal))
                {
                    score += 3m;
                    result.Details.Add("Cong thuc Inventory Notice dung chinh xac.");
                }
                else
                {
                    result.Errors.Add($"Cong thuc Inventory Notice chua dung. Hien tai: '{noticeColumn.CalculatedColumnFormula}'.");
                }

                var hasStrictText = normalizedFormula.Contains("\"LOW\",\"\"", StringComparison.Ordinal);
                if (hasStrictText)
                {
                    score += 1m;
                    result.Details.Add("Noi dung text trong IF dung chinh ta: 'Low' va rong.");
                }
                else
                {
                    result.Errors.Add("Text trong cong thuc IF chua dung ('Low' / \"\").");
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

    public class P12T6Grader : ITaskGrader
    {
        public string TaskId => "P12-T6";
        public string TaskName => "Inventory chart: title + data labels outEnd";
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
                    result.Errors.Add("Khong tim thay chart Inventory Level % can kiem tra.");
                    return result;
                }

                score += 1m;
                result.Details.Add("Tim thay dung chart Inventory Level %.");

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
                    result.Details.Add("Data labels dung vi tri ben phai cot (outEnd).");
                }
                else
                {
                    result.Errors.Add($"Vi tri data label chua dung. Hien tai: '{dLblPos}'.");
                }

                var showVal = xml.SelectSingleNode("//c:ser/c:dLbls/c:showVal", ns)?.Attributes?["val"]?.Value ?? string.Empty;
                if (string.Equals(showVal, "1", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Data labels da hien thi gia tri.");
                }
                else
                {
                    result.Errors.Add("Data labels chua hien thi gia tri.");
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
