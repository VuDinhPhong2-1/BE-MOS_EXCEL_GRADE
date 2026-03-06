using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Xml;

namespace MOS.ExcelGrading.Core.Graders.Project02
{
    public class P02T7Grader : ITaskGrader
    {
        public string TaskId => "P02-T7";
        public string TaskName => "Ap dung chart style Colorful Palette 2";
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
                var ws = studentSheet.Workbook.Worksheets["New Policy"];
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'New Policy'");
                    return result;
                }

                var chart = ws.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Khong tim thay bieu do tren New Policy");
                    return result;
                }

                var xml = chart.ChartXml;
                var ns = new XmlNamespaceManager(xml.NameTable);
                ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
                ns.AddNamespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                ns.AddNamespace("c14", "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
                ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                decimal score = 0;
                score += 1m; // tim thay chart

                // Colorful Palette 2 trong file dap an duoc luu voi color style id=11.
                // Check id nay de phan biet chac chan voi Colorful Palette 1.
                var colorStyleIdRaw = chart.StyleManager?.ColorsXml?.DocumentElement?.Attributes?["id"]?.Value?.Trim();
                var isPalette2 = string.Equals(colorStyleIdRaw, "11", StringComparison.Ordinal);
                if (isPalette2)
                {
                    score += 1m;
                    result.Details.Add("Chart Color Style ID = 11 (Colorful Palette 2)");
                }
                else
                {
                    result.Errors.Add($"Chart Color Style ID khong dung Palette 2. Hien tai: '{colorStyleIdRaw ?? "null"}' (mong doi: '11').");
                }

                var seriesNodes = xml.SelectNodes("//c:barChart/c:ser", ns);
                if (seriesNodes == null || seriesNodes.Count == 0)
                {
                    result.Errors.Add("Khong doc duoc danh sach series de kiem tra palette");
                    result.Score = score;
                    return result;
                }

                var signatures = new List<(string Color, string LumMod)>();
                foreach (XmlNode ser in seriesNodes)
                {
                    var clrNode = ser.SelectSingleNode("c:spPr/a:solidFill/a:schemeClr", ns);
                    var color = clrNode?.Attributes?["val"]?.Value?.Trim() ?? string.Empty;
                    var lumMod = clrNode?.SelectSingleNode("a:lumMod", ns)?.Attributes?["val"]?.Value?.Trim() ?? string.Empty;
                    signatures.Add((color, lumMod));
                }

                var expected = new List<(string Color, string LumMod)>
                {
                    ("accent1", ""),
                    ("accent3", ""),
                    ("accent5", ""),
                    ("accent1", "60000"),
                    ("accent3", "60000"),
                    ("accent5", "60000"),
                };

                var compareCount = Math.Min(expected.Count, signatures.Count);
                var primaryMatch = 0;
                var secondaryMatch = 0;

                for (var i = 0; i < compareCount; i++)
                {
                    var actual = signatures[i];
                    var exp = expected[i];
                    var colorMatch = string.Equals(actual.Color, exp.Color, StringComparison.OrdinalIgnoreCase);
                    var lumMatch = string.Equals(actual.LumMod, exp.LumMod, StringComparison.OrdinalIgnoreCase);

                    if (i < 3)
                    {
                        if (colorMatch && lumMatch)
                        {
                            primaryMatch++;
                        }
                    }
                    else
                    {
                        if (colorMatch && lumMatch)
                        {
                            secondaryMatch++;
                        }
                    }
                }

                if (primaryMatch == 3)
                {
                    score += 1m;
                    result.Details.Add("3 series dau dung mau cua Colorful Palette 2 (accent1/accent3/accent5)");
                }
                else
                {
                    result.Errors.Add($"3 series dau chua dung Palette 2 ({primaryMatch}/3)");
                }

                if (secondaryMatch == 3)
                {
                    score += 1m;
                    result.Details.Add("3 series tiep theo dung tone Palette 2 (lumMod=60000)");
                }
                else
                {
                    result.Errors.Add($"3 series tone mau tiep theo chua dung Palette 2 ({secondaryMatch}/3)");
                }

                var signatureText = string.Join(" | ", signatures.Select((s, idx) =>
                    $"S{idx + 1}:{s.Color}{(string.IsNullOrWhiteSpace(s.LumMod) ? "" : $"(lum={s.LumMod})")}"));
                result.Details.Add($"Palette signature: {signatureText}");

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
