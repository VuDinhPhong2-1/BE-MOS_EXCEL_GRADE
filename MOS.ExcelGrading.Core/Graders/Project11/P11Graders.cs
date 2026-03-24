using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project11
{
    public class P11T1Grader : ITaskGrader
    {
        public string TaskId => "P11-T1";
        public string TaskName => "Games: merge A12:B12 through A18:B18";
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
                var ws = P11GraderHelpers.GetSheet(studentSheet.Workbook, "Games");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Games'.");
                    return result;
                }

                decimal score = 0m;
                var expectedRanges = Enumerable.Range(12, 7)
                    .Select(row => $"A{row}:B{row}")
                    .ToList();

                var matchedCount = expectedRanges.Count(range => P11GraderHelpers.HasMergeRange(ws, range));
                if (matchedCount == expectedRanges.Count)
                {
                    score += 2m;
                    result.Details.Add("Da merge day du cac vung A12:B12 den A18:B18.");
                }
                else
                {
                    var missing = expectedRanges.Where(range => !P11GraderHelpers.HasMergeRange(ws, range));
                    result.Errors.Add($"Thieu merge ranges: {string.Join(", ", missing)}.");
                }

                var centeredCount = expectedRanges.Count(range =>
                {
                    var topLeft = range.Split(':')[0];
                    return P11GraderHelpers.IsMergedCellCentered(ws, topLeft);
                });
                if (centeredCount == expectedRanges.Count)
                {
                    score += 1m;
                    result.Details.Add("Cac o merge duoc can giua ngang/doc dung yeu cau.");
                }
                else
                {
                    result.Errors.Add($"Can giua merge chua dung ({centeredCount}/{expectedRanges.Count}).");
                }

                var hasOnlyExpectedMerges = ws.MergedCells.Count == expectedRanges.Count;
                if (hasOnlyExpectedMerges)
                {
                    score += 1m;
                    result.Details.Add("So luong merge range dung (7).");
                }
                else
                {
                    result.Errors.Add($"So luong merge range chua dung. Hien tai: {ws.MergedCells.Count}.");
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

    public class P11T2Grader : ITaskGrader
    {
        public string TaskId => "P11-T2";
        public string TaskName => "Shareholders Info: row height for Annual Report row = 30";
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
                var ws = P11GraderHelpers.GetSheet(studentSheet.Workbook, "Shareholders Info");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Shareholders Info'.");
                    return result;
                }

                decimal score = 0m;
                var annualRowIndex = -1;
                for (var row = 1; row <= ws.Dimension.End.Row; row++)
                {
                    var rowText = ws.Cells[row, 1, row, ws.Dimension.End.Column].Text;
                    if (!string.IsNullOrWhiteSpace(rowText)
                        && rowText.Contains("Annual Report", StringComparison.OrdinalIgnoreCase))
                    {
                        annualRowIndex = row;
                        break;
                    }
                }

                if (annualRowIndex < 0)
                {
                    result.Errors.Add("Khong tim thay dong chua gia tri 'Annual Report'.");
                    return result;
                }

                score += 1m;
                result.Details.Add($"Tim thay dong 'Annual Report' tai hang {annualRowIndex}.");

                var targetRow = ws.Row(annualRowIndex);
                if (Math.Abs(targetRow.Height - 30d) <= 0.1d)
                {
                    score += 3m;
                    result.Details.Add("Do cao dong dung 30.");
                }
                else
                {
                    result.Errors.Add($"Do cao dong chua dung. Hien tai: {targetRow.Height:0.##}, mong doi: 30.");
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

    public class P11T3Grader : ITaskGrader
    {
        public string TaskId => "P11-T3";
        public string TaskName => "Rename sheet Outdoor Toys to Outdoor Sports";
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
                decimal score = 0m;

                var hasOutdoorSports = P11GraderHelpers.GetSheet(workbook, "Outdoor Sports") != null;
                var hasOutdoorToys = P11GraderHelpers.GetSheet(workbook, "Outdoor Toys") != null;

                if (hasOutdoorSports)
                {
                    score += 2m;
                    result.Details.Add("Da ton tai sheet moi 'Outdoor Sports'.");
                }
                else
                {
                    result.Errors.Add("Khong tim thay sheet 'Outdoor Sports'.");
                }

                if (!hasOutdoorToys)
                {
                    score += 2m;
                    result.Details.Add("Khong con sheet cu 'Outdoor Toys'.");
                }
                else
                {
                    result.Errors.Add("Van con sheet cu 'Outdoor Toys'.");
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

    public class P11T4Grader : ITaskGrader
    {
        public string TaskId => "P11-T4";
        public string TaskName => "Shareholders Info: C5 hyperlink and display text";
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
                var ws = P11GraderHelpers.GetSheet(studentSheet.Workbook, "Shareholders Info");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Shareholders Info'.");
                    return result;
                }

                decimal score = 0m;
                var hyperlink = ws.Cells["C5"].Hyperlink;
                if (hyperlink != null)
                {
                    score += 1m;
                    result.Details.Add("C5 da co hyperlink.");
                }
                else
                {
                    result.Errors.Add("C5 chua co hyperlink.");
                    result.Score = score;
                    return result;
                }

                var linkText = hyperlink.OriginalString ?? hyperlink.ToString() ?? string.Empty;
                var normalized = P11GraderHelpers.NormalizeUrl(linkText);
                if (string.Equals(normalized, "http://tailspintoys.com/beyond.html", StringComparison.Ordinal))
                {
                    score += 2m;
                    result.Details.Add("Hyperlink dung URL yeu cau.");
                }
                else
                {
                    result.Errors.Add($"URL hyperlink chua dung. Hien tai: '{linkText}'.");
                }

                var displayText = ws.Cells["C5"].Text ?? string.Empty;
                if (string.Equals(displayText, "More Info", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Van ban hien thi o C5 dung chinh ta: 'More Info'.");
                }
                else
                {
                    result.Errors.Add($"Van ban hien thi C5 chua dung. Hien tai: '{displayText}'.");
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

    public class P11T5Grader : ITaskGrader
    {
        public string TaskId => "P11-T5";
        public string TaskName => "Set print scaling per worksheet (fit to one page)";
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
                var worksheets = studentSheet.Workbook.Worksheets
                    .Where(ws => ws is not OfficeOpenXml.ExcelChartsheet)
                    .ToList();
                if (worksheets.Count == 0)
                {
                    result.Errors.Add("Khong tim thay worksheet de kiem tra print scaling.");
                    return result;
                }

                decimal score = 0m;
                var fitToPageCount = 0;
                foreach (var ws in worksheets)
                {
                    var ns = P11GraderHelpers.CreateWorkbookNamespaceManager(ws.WorksheetXml);
                    var pageSetUpPr = ws.WorksheetXml.SelectSingleNode("//x:sheetPr/x:pageSetUpPr", ns);
                    var fitToPage = pageSetUpPr?.Attributes?["fitToPage"]?.Value ?? string.Empty;
                    if (string.Equals(fitToPage, "1", StringComparison.Ordinal))
                    {
                        fitToPageCount++;
                    }
                }

                if (fitToPageCount == worksheets.Count)
                {
                    score += 3m;
                    result.Details.Add($"Da bat fitToPage cho tat ca worksheet ({fitToPageCount}/{worksheets.Count}).");
                }
                else
                {
                    result.Errors.Add($"Chua bat fitToPage day du ({fitToPageCount}/{worksheets.Count}).");
                }

                if (worksheets.Count == 4)
                {
                    score += 1m;
                    result.Details.Add("So worksheet dung 4 trang can cau hinh in.");
                }
                else
                {
                    result.Errors.Add($"So worksheet hien tai la {worksheets.Count}, mong doi 4.");
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

    public class P11T6Grader : ITaskGrader
    {
        public string TaskId => "P11-T6";
        public string TaskName => "Costs: print titles rows 1:3";
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

                decimal score = 0m;
                var ns = P11GraderHelpers.CreateWorkbookNamespaceManager(workbookXml);
                var printTitleNode = workbookXml.SelectSingleNode(
                    "//x:definedNames/x:definedName[@name='_xlnm.Print_Titles']",
                    ns);
                if (printTitleNode == null)
                {
                    result.Errors.Add("Khong tim thay defined name _xlnm.Print_Titles.");
                    return result;
                }

                score += 1m;
                result.Details.Add("Tim thay defined name _xlnm.Print_Titles.");

                var localSheetIdText = printTitleNode.Attributes?["localSheetId"]?.Value ?? string.Empty;
                var costsSheetIndex = P11GraderHelpers.GetSheetIndex0Based(workbook, "Costs");
                if (int.TryParse(localSheetIdText, out var localSheetId)
                    && costsSheetIndex >= 0
                    && localSheetId == costsSheetIndex)
                {
                    score += 1m;
                    result.Details.Add("Print_Titles duoc dat dung tren sheet Costs.");
                }
                else
                {
                    result.Errors.Add($"localSheetId chua dung. Hien tai: '{localSheetIdText}', mong doi: {costsSheetIndex}.");
                }

                var normalizedValue = (printTitleNode.InnerText ?? string.Empty)
                    .Replace("$", string.Empty, StringComparison.Ordinal)
                    .Replace("'", string.Empty, StringComparison.Ordinal)
                    .Replace(" ", string.Empty, StringComparison.Ordinal)
                    .ToUpperInvariant();
                if (string.Equals(normalizedValue, "COSTS!1:3", StringComparison.Ordinal))
                {
                    score += 2m;
                    result.Details.Add("Gia tri Print Titles dung: Costs!$1:$3.");
                }
                else
                {
                    result.Errors.Add($"Gia tri Print Titles chua dung. Hien tai: '{printTitleNode.InnerText}'.");
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
