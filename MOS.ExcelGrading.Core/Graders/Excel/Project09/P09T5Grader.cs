using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project09
{
    public class P09T5Grader : ITaskGrader
    {
        public string TaskId => "P09-T5";
        public string TaskName => "Subtotal theo Shirt Color, ngat trang va Grand Total";
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
                var ws = P09GraderHelpers.GetSheet(studentSheet.Workbook, "Shirt Orders");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Shirt Orders'.");
                    return result;
                }

                decimal score = 0m;
                var ns = P09GraderHelpers.CreateWorksheetNamespaceManager(ws.WorksheetXml);

                var sheetFormatNode = ws.WorksheetXml.SelectSingleNode("//x:sheetFormatPr", ns);
                var outlineLevelText = sheetFormatNode?.Attributes?["outlineLevelRow"]?.Value ?? string.Empty;
                if (int.TryParse(outlineLevelText, out var outlineLevel) && outlineLevel >= 1)
                {
                    score += 1.5m;
                    result.Details.Add($"sheetFormatPr hop le outlineLevelRow={outlineLevel}.");
                }
                else
                {
                    var subtotalCells = new[] { "F46", "F103", "F156", "F200" };
                    var hasSubtotalStructure = subtotalCells.All(cell =>
                    {
                        var formula = P09GraderHelpers.NormalizeFormula(ws.Cells[cell].Formula);
                        return formula.StartsWith("SUBTOTAL(", StringComparison.Ordinal)
                               || formula.StartsWith("SUM(", StringComparison.Ordinal);
                    });

                    if (hasSubtotalStructure)
                    {
                        score += 1.5m;
                        result.Details.Add("Không có outlineLevelRow ro rang nhung da co cau truc subtotal hop le.");
                    }
                    else
                    {
                        result.Errors.Add($"outlineLevelRow/cau truc subtotal chưa đúng. Hiện tại outlineLevelRow='{outlineLevelText}'.");
                    }
                }

                var detectedBreakIds = new HashSet<int>();
                var scanEndRow = Math.Max(ws.Dimension?.End.Row ?? 0, 300);
                for (var row = 1; row <= scanEndRow; row++)
                {
                    if (ws.Row(row).PageBreak)
                    {
                        detectedBreakIds.Add(row);
                    }
                }

                // EPPlus may not keep rowBreaks nodes in WorksheetXml after load, so keep XML as fallback only.
                if (detectedBreakIds.Count == 0)
                {
                    var rowBreaksNode = ws.WorksheetXml.SelectSingleNode("//x:rowBreaks", ns);
                    var manualBreakNodes = rowBreaksNode?.SelectNodes("x:brk[@man='1']", ns);
                    var allBreakNodes = rowBreaksNode?.SelectNodes("x:brk", ns);
                    var breakNodes = manualBreakNodes ?? allBreakNodes;
                    if (breakNodes != null)
                    {
                        foreach (XmlNode breakNode in breakNodes)
                        {
                            var idText = breakNode.Attributes?["id"]?.Value ?? string.Empty;
                            if (int.TryParse(idText, out var id))
                            {
                                detectedBreakIds.Add(id);
                            }
                        }
                    }
                }

                var detectedManualBreakCount = detectedBreakIds.Count;

                if (detectedManualBreakCount is 3 or 4)
                {
                    score += 1m;
                    result.Details.Add("Số lượng manual row break hop le (3 hoac 4).");
                }
                else
                {
                    result.Errors.Add(
                        $"Số lượng manual row break chưa đúng. detected={detectedManualBreakCount}, mong đợi 3 hoac 4.");
                }

                var actualBreakIds = detectedBreakIds;

                var requiredBreakAnchors = new[] { 46, 103, 156 };
                var optionalFinalBreakAnchor = 201;
                var tolerance = 1;

                var hasRequiredBreakAnchors = requiredBreakAnchors.All(expected =>
                    actualBreakIds.Any(actual => Math.Abs(actual - expected) <= tolerance));

                var hasOnlyAcceptedBreaks = actualBreakIds.All(actual =>
                    requiredBreakAnchors.Any(expected => Math.Abs(actual - expected) <= tolerance)
                    || Math.Abs(actual - optionalFinalBreakAnchor) <= tolerance);

                if (hasRequiredBreakAnchors
                    && hasOnlyAcceptedBreaks
                    && actualBreakIds.Count is >= 3 and <= 4)
                {
                    score += 1.5m;
                    result.Details.Add("Danh sach row break id hop le (cac moc 46, 103, 156; co the co them moc 201).");
                }
                else
                {
                    result.Errors.Add(
                        $"Row break id chưa đúng. Hiện tại: {string.Join(", ", actualBreakIds.OrderBy(x => x))}. Mong doi quanh cac moc 46, 103, 156 (va tuy chon 201).");
                }

                var expectedFormulas = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
                {
                    ["F46"] = new[] { "SUBTOTAL(9,F6:F45)", "SUM(F6:F45)" },
                    ["F103"] = new[] { "SUBTOTAL(9,F47:F102)", "SUM(F47:F102)" },
                    ["F156"] = new[] { "SUBTOTAL(9,F104:F155)", "SUM(F104:F155)" },
                    ["F200"] = new[] { "SUBTOTAL(9,F157:F199)", "SUM(F157:F199)" },
                    ["F201"] = new[] { "SUBTOTAL(9,F6:F199)", "SUM(F46,F103,F156,F200)", "SUM(F6:F200)" }
                };

                foreach (var pair in expectedFormulas)
                {
                    var actualFormula = P09GraderHelpers.NormalizeFormula(ws.Cells[pair.Key].Formula);
                    var hasAcceptedFormula = pair.Value.Any(acceptedFormula =>
                        string.Equals(
                            actualFormula,
                            P09GraderHelpers.NormalizeFormula(acceptedFormula),
                            StringComparison.Ordinal));

                    if (hasAcceptedFormula)
                    {
                        score += 0.8m;
                        result.Details.Add($"{pair.Key} dung cong thuc.");
                    }
                    else
                    {
                        result.Errors.Add(
                            $"{pair.Key} sai cong thuc. Hiện tại: '{ws.Cells[pair.Key].Formula}'.");
                    }
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

// minor-sync: non-functional graders update



