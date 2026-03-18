using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

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
                var ws = P09GraderHelpers.GetSheet(studentSheet.Workbook, "Shirt Orders");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Shirt Orders'.");
                    return result;
                }

                decimal score = 0m;
                var ns = P09GraderHelpers.CreateWorksheetNamespaceManager(ws.WorksheetXml);

                var sheetFormatNode = ws.WorksheetXml.SelectSingleNode("//x:sheetFormatPr", ns);
                var outlineLevelText = sheetFormatNode?.Attributes?["outlineLevelRow"]?.Value ?? string.Empty;
                if (int.TryParse(outlineLevelText, out var outlineLevel) && outlineLevel == 2)
                {
                    score += 1.5m;
                    result.Details.Add("sheetFormatPr dung outlineLevelRow=2.");
                }
                else
                {
                    result.Errors.Add($"outlineLevelRow chua dung. Hien tai: '{outlineLevelText}'.");
                }

                var rowBreaksNode = ws.WorksheetXml.SelectSingleNode("//x:rowBreaks", ns);
                var manualBreakCountText = rowBreaksNode?.Attributes?["manualBreakCount"]?.Value ?? string.Empty;
                if (int.TryParse(manualBreakCountText, out var manualBreakCount) && manualBreakCount == 4)
                {
                    score += 1m;
                    result.Details.Add("manualBreakCount dung = 4.");
                }
                else
                {
                    result.Errors.Add($"manualBreakCount chua dung. Hien tai: '{manualBreakCountText}'.");
                }

                var expectedBreakIds = new HashSet<int> { 46, 103, 156, 201 };
                var breakNodes = rowBreaksNode?.SelectNodes("x:brk", ns);
                var actualBreakIds = new HashSet<int>();
                if (breakNodes != null)
                {
                    foreach (XmlNode breakNode in breakNodes)
                    {
                        var idText = breakNode.Attributes?["id"]?.Value ?? string.Empty;
                        if (int.TryParse(idText, out var id))
                        {
                            actualBreakIds.Add(id);
                        }
                    }
                }

                if (expectedBreakIds.SetEquals(actualBreakIds))
                {
                    score += 1.5m;
                    result.Details.Add("Danh sach row break id dung: 46, 103, 156, 201.");
                }
                else
                {
                    result.Errors.Add(
                        $"Row break id chua dung. Hien tai: {string.Join(", ", actualBreakIds.OrderBy(x => x))}.");
                }

                var expectedFormulas = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["F46"] = "SUBTOTAL(9,F6:F45)",
                    ["F103"] = "SUBTOTAL(9,F47:F102)",
                    ["F156"] = "SUBTOTAL(9,F104:F155)",
                    ["F200"] = "SUBTOTAL(9,F157:F199)",
                    ["F201"] = "SUBTOTAL(9,F6:F199)"
                };

                foreach (var pair in expectedFormulas)
                {
                    var actualFormula = P09GraderHelpers.NormalizeFormula(ws.Cells[pair.Key].Formula);
                    var expectedFormula = P09GraderHelpers.NormalizeFormula(pair.Value);
                    if (string.Equals(actualFormula, expectedFormula, StringComparison.Ordinal))
                    {
                        score += 0.8m;
                        result.Details.Add($"{pair.Key} dung cong thuc SUBTOTAL.");
                    }
                    else
                    {
                        result.Errors.Add(
                            $"{pair.Key} sai cong thuc. Hien tai: '{ws.Cells[pair.Key].Formula}'.");
                    }
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
