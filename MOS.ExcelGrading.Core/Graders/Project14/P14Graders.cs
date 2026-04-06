using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project14
{
    internal static class P14GraderHelpers
    {
        public static ExcelWorksheet? GetSheet(ExcelWorkbook workbook, string name)
        {
            return workbook.Worksheets.FirstOrDefault(sheet =>
                string.Equals((sheet.Name ?? string.Empty).Trim(), name.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        public static string NormalizeFormula(string? formula)
        {
            var normalized = (formula ?? string.Empty)
                .Trim()
                .Replace("=", string.Empty, StringComparison.Ordinal)
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Replace(" ", string.Empty, StringComparison.Ordinal)
                .Replace("_xlfn.", string.Empty, StringComparison.OrdinalIgnoreCase)
                .Replace(";", ",", StringComparison.Ordinal)
                .Replace("@", string.Empty, StringComparison.Ordinal)
                .ToUpperInvariant();

            // Excel may store table formulas in equivalent forms:
            // Table1[[#This Row],[Col]] or Table1[Col].
            return normalized
                .Replace("[[#THISROW],[", "[", StringComparison.Ordinal)
                .Replace("]]", "]", StringComparison.Ordinal);
        }

        public static string NormalizePrintArea(string? value)
        {
            return (value ?? string.Empty)
                .Trim()
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Replace("'", string.Empty, StringComparison.Ordinal)
                .Replace(" ", string.Empty, StringComparison.Ordinal)
                .ToUpperInvariant();
        }

        public static string GetSingleFilterValue(ExcelTable table, int colId)
        {
            var ns = new XmlNamespaceManager(table.TableXml.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var node = table.TableXml.SelectSingleNode(
                $"//x:autoFilter/x:filterColumn[@colId='{colId}']/x:filters/x:filter",
                ns);
            return node?.Attributes?["val"]?.Value?.Trim() ?? string.Empty;
        }

        public static ExcelTable? FindTableByAddress(ExcelWorksheet worksheet, string expectedAddress)
        {
            return worksheet.Tables.FirstOrDefault(table =>
                string.Equals(
                    NormalizePrintArea(table.Address.Address),
                    NormalizePrintArea(expectedAddress),
                    StringComparison.OrdinalIgnoreCase));
        }
    }

    public class P14T1Grader : ITaskGrader
    {
        public string TaskId => "P14-T1";
        public string TaskName => "January: set print area A4:F20";
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
                var workbook = studentSheet.Workbook;
                var workbookXml = workbook.WorkbookXml;
                if (workbookXml == null)
                {
                    result.Errors.Add("Khong doc duoc workbook XML.");
                    return result;
                }

                var januaryIndex = workbook.Worksheets.ToList()
                    .FindIndex(ws => string.Equals(ws.Name, "January", StringComparison.OrdinalIgnoreCase));
                if (januaryIndex < 0)
                {
                    result.Errors.Add("Khong tim thay sheet 'January'.");
                    return result;
                }

                decimal score = 0m;
                var ns = new XmlNamespaceManager(workbookXml.NameTable);
                ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var printAreaNode = workbookXml.SelectSingleNode(
                    $"//x:definedNames/x:definedName[@name='_xlnm.Print_Area' and @localSheetId='{januaryIndex}']",
                    ns);

                if (printAreaNode != null)
                {
                    score += 2m;
                    result.Details.Add("Tim thay Print_Area cho sheet January.");
                }
                else
                {
                    result.Errors.Add("Khong tim thay Print_Area cho sheet January.");
                    result.Score = score;
                    return result;
                }

                var normalized = P14GraderHelpers.NormalizePrintArea(printAreaNode.InnerText);
                var isExpected = string.Equals(normalized, "JANUARY!A4:F20", StringComparison.Ordinal)
                                 || string.Equals(normalized, "TABLE4[[#ALL],[CLIENTID]:[POLICYTYPE]]", StringComparison.Ordinal);
                if (isExpected)
                {
                    score += 2m;
                    result.Details.Add("Print_Area dung pham vi yeu cau A4:F20.");
                }
                else
                {
                    result.Errors.Add($"Gia tri Print_Area chua dung. Hien tai: '{printAreaNode.InnerText}'.");
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

    public class P14T2Grader : ITaskGrader
    {
        public string TaskId => "P14-T2";
        public string TaskName => "March: filter Policy Type = PM";
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
                var ws = P14GraderHelpers.GetSheet(studentSheet.Workbook, "March");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'March'.");
                    return result;
                }

                decimal score = 0m;
                var table = P14GraderHelpers.FindTableByAddress(ws, "A4:G24");
                if (table != null)
                {
                    score += 1m;
                    result.Details.Add("Tim thay table March A4:G24.");
                }
                else
                {
                    result.Errors.Add("Khong tim thay table March A4:G24.");
                    result.Score = score;
                    return result;
                }

                var filterValue = P14GraderHelpers.GetSingleFilterValue(table, 5);
                if (string.Equals(filterValue, "PM", StringComparison.Ordinal))
                {
                    score += 3m;
                    result.Details.Add("Filter cot Policy Type dung gia tri 'PM'.");
                }
                else
                {
                    result.Errors.Add($"Filter Policy Type chua dung. Hien tai: '{filterValue}'.");
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

    public class P14T3Grader : ITaskGrader
    {
        public string TaskId => "P14-T3";
        public string TaskName => "February: Discount formula";
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
                var ws = P14GraderHelpers.GetSheet(studentSheet.Workbook, "February");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'February'.");
                    return result;
                }

                var table = P14GraderHelpers.FindTableByAddress(ws, "A4:G18");
                if (table == null)
                {
                    result.Errors.Add("Khong tim thay table February A4:G18.");
                    return result;
                }

                var column = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name, "Discount", StringComparison.OrdinalIgnoreCase));
                if (column == null)
                {
                    result.Errors.Add("Khong tim thay cot 'Discount'.");
                    return result;
                }

                var actual = P14GraderHelpers.NormalizeFormula(column.CalculatedColumnFormula);
                var expected = P14GraderHelpers.NormalizeFormula("IF(Table1[[#This Row],[Years as Member]]>3,\"Yes\",\"No\")");
                if (string.Equals(actual, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Cong thuc cot Discount dung chinh xac.");
                }
                else
                {
                    result.Errors.Add($"Cong thuc Discount chua dung. Hien tai: '{column.CalculatedColumnFormula}'.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }

    public class P14T4Grader : ITaskGrader
    {
        public string TaskId => "P14-T4";
        public string TaskName => "February: Policy Type formula LEFT(...,2)";
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
                var ws = P14GraderHelpers.GetSheet(studentSheet.Workbook, "February");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'February'.");
                    return result;
                }

                var table = P14GraderHelpers.FindTableByAddress(ws, "A4:G18");
                if (table == null)
                {
                    result.Errors.Add("Khong tim thay table February A4:G18.");
                    return result;
                }

                var column = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name, "Policy Type", StringComparison.OrdinalIgnoreCase));
                if (column == null)
                {
                    result.Errors.Add("Khong tim thay cot 'Policy Type'.");
                    return result;
                }

                var actual = P14GraderHelpers.NormalizeFormula(column.CalculatedColumnFormula);
                var expected = P14GraderHelpers.NormalizeFormula("LEFT(Table1[[#This Row],[Policy Number]],2)");
                if (string.Equals(actual, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Cong thuc cot Policy Type dung chinh xac.");
                }
                else
                {
                    result.Errors.Add($"Cong thuc Policy Type chua dung. Hien tai: '{column.CalculatedColumnFormula}'.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }

    public class P14T5Grader : ITaskGrader
    {
        public string TaskId => "P14-T5";
        public string TaskName => "Summary: chart alt text = Renewal Data";
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
                var ws = P14GraderHelpers.GetSheet(studentSheet.Workbook, "Summary");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Summary'.");
                    return result;
                }

                var chart = ws.Drawings.FirstOrDefault(d => d is OfficeOpenXml.Drawing.Chart.ExcelChart);
                if (chart == null)
                {
                    result.Errors.Add("Khong tim thay chart tren sheet 'Summary'.");
                    return result;
                }

                var description = chart.GetType().GetProperty("Description")?.GetValue(chart)?.ToString() ?? string.Empty;
                if (string.Equals(description.Trim(), "Renewal Data", StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Alt text chart dung: 'Renewal Data'.");
                }
                else
                {
                    result.Errors.Add($"Alt text chart chua dung. Hien tai: '{description}'.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }

    public class P14T6Grader : ITaskGrader
    {
        public string TaskId => "P14-T6";
        public string TaskName => "Sales: Auction ID formula RANDBETWEEN(1000,2000)";
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
                var ws = P14GraderHelpers.GetSheet(studentSheet.Workbook, "Sales");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Sales'.");
                    return result;
                }

                var table = P14GraderHelpers.FindTableByAddress(ws, "A3:F17");
                if (table == null)
                {
                    result.Errors.Add("Khong tim thay table Sales A3:F17.");
                    return result;
                }

                var column = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name, "Auction ID", StringComparison.OrdinalIgnoreCase));
                if (column == null)
                {
                    result.Errors.Add("Khong tim thay cot 'Auction ID'.");
                    return result;
                }

                var actual = P14GraderHelpers.NormalizeFormula(column.CalculatedColumnFormula);
                var expected = P14GraderHelpers.NormalizeFormula("RANDBETWEEN(1000,2000)");
                if (string.Equals(actual, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Cong thuc Auction ID dung chinh xac.");
                }
                else
                {
                    result.Errors.Add($"Cong thuc Auction ID chua dung. Hien tai: '{column.CalculatedColumnFormula}'.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }
}
