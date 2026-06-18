using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System.Xml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project04
{
    public class P04T6Grader : ITaskGrader
    {
        public string TaskId => "P04-T6";
        public string TaskName => "Ð?t tiêu d? tr?c d?c chính là 'Hours'";
        public decimal MaxScore => 4;

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
                var ws = P04GraderHelpers.GetSheet(studentSheet, "Number of course hours");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet Number of course hours");
                    return result;
                }

                var charts = ws.Drawings.OfType<ExcelChart>().ToList();
                if (charts.Count == 0)
                {
                    result.Errors.Add("Không tìm th?y chart trên Number of course hours");
                    return result;
                }

                result.Score += 1m;
                result.Details.Add("Tim th?y chart c?n ch?m");

                var chartInfo = charts
                    .Select(c => new { Chart = c, YTitle = GetPrimaryVerticalAxisTitle(c, ws) })
                    .OrderByDescending(x => GetTitleCandidateScore(x.YTitle))
                    .First();

                var yTitle = chartInfo.YTitle;
                if (!string.IsNullOrWhiteSpace(yTitle))
                {
                    result.Score += 1.5m;
                    result.Details.Add($"Tr?c d?c dã có tiêu d?: '{yTitle}'");
                }
                else
                {
                    result.Errors.Add("Tr?c d?c chua có tiêu d?");
                }

                // Compare exact text; leading/trailing spaces must be treated as incorrect.
                if (string.Equals(yTitle, "Hours", StringComparison.Ordinal))
                {
                    result.Score += 1.5m;
                    result.Details.Add("Tiêu d? tr?c d?c dúng 'Hours'");
                }
                else
                {
                    result.Errors.Add($"Tiêu d? tr?c d?c chua dúng chính xác. Hi?n t?i: '{yTitle}', mong d?i 'Hours'.");
                }

                result.Score = Math.Min(MaxScore, result.Score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }

        private static int GetTitleCandidateScore(string? title)
        {
            var value = title ?? string.Empty;
            if (string.Equals(value, "Hours", StringComparison.Ordinal))
            {
                return 2;
            }

            return string.IsNullOrWhiteSpace(value) ? 0 : 1;
        }

        private static string GetPrimaryVerticalAxisTitle(ExcelChart chart, ExcelWorksheet contextSheet)
        {
            var directText = chart.YAxis?.Title?.Text;
            if (!string.IsNullOrEmpty(directText) && !string.IsNullOrWhiteSpace(directText))
            {
                return directText;
            }

            var xml = chart.ChartXml;
            var ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            // Prefer primary vertical value axis title first, then fallback to any axis title.
            var axisTitleNodes = xml.SelectNodes("//c:plotArea/c:valAx[c:axPos[@val='l']]/c:title", ns);
            if (axisTitleNodes == null || axisTitleNodes.Count == 0)
            {
                axisTitleNodes = xml.SelectNodes("//c:plotArea/*[self::c:valAx or self::c:catAx or self::c:dateAx or self::c:serAx]/c:title", ns);
            }
            if (axisTitleNodes == null || axisTitleNodes.Count == 0)
            {
                return directText ?? string.Empty;
            }

            foreach (XmlNode titleNode in axisTitleNodes)
            {
                var richTextNodes = titleNode.SelectNodes(".//a:t", ns);
                if (richTextNodes != null && richTextNodes.Count > 0)
                {
                    var richText = string.Concat(richTextNodes.Cast<XmlNode>().Select(n => n.InnerText));
                    if (!string.IsNullOrWhiteSpace(richText))
                    {
                        return richText;
                    }
                }

                var cachedTextNodes = titleNode.SelectNodes(".//c:strRef//c:v", ns);
                if (cachedTextNodes != null && cachedTextNodes.Count > 0)
                {
                    var cachedText = string.Concat(cachedTextNodes.Cast<XmlNode>().Select(n => n.InnerText));
                    if (!string.IsNullOrWhiteSpace(cachedText))
                    {
                        return cachedText;
                    }
                }

                var formulaNode = titleNode.SelectSingleNode(".//c:strRef/c:f", ns);
                if (formulaNode != null && !string.IsNullOrWhiteSpace(formulaNode.InnerText))
                {
                    var linkedText = TryResolveLinkedTitle(contextSheet, formulaNode.InnerText);
                    if (!string.IsNullOrWhiteSpace(linkedText))
                    {
                        return linkedText;
                    }
                }
            }

            return directText ?? string.Empty;
        }

        private static string TryResolveLinkedTitle(ExcelWorksheet contextSheet, string formulaText)
        {
            var formula = (formulaText ?? string.Empty).Trim();
            if (formula.Length == 0)
            {
                return string.Empty;
            }

            var ws = contextSheet;
            var address = formula;

            var excl = formula.LastIndexOf('!');
            if (excl >= 0)
            {
                var rawSheet = formula[..excl].Trim().Trim('\'');
                address = formula[(excl + 1)..].Trim();

                var targetSheet = contextSheet.Workbook.Worksheets.FirstOrDefault(s =>
                    string.Equals(s.Name, rawSheet, StringComparison.OrdinalIgnoreCase));
                if (targetSheet != null)
                {
                    ws = targetSheet;
                }
            }

            address = address.Replace("$", string.Empty, StringComparison.Ordinal);
            if (!ExcelCellBase.IsValidAddress(address))
            {
                return string.Empty;
            }

            return ws.Cells[address].Text ?? string.Empty;
        }
    }
}

