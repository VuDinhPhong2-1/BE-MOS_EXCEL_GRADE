using System.Globalization;
using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project07
{
    public class P07T4Grader : ITaskGrader
    {
        public string TaskId => "P07-T4";
        public string TaskName => "Total Cookie Sales A3:A8 dung Pattern Fill theo mau de bai";
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
                var ws = P07GraderHelpers.GetSheet(studentSheet, "Total Cookie Sales");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Total Cookie Sales'.");
                    return result;
                }

                var stylesXml = ws.Workbook.StylesXml;
                var ns = new XmlNamespaceManager(stylesXml.NameTable);
                ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                const int startRow = 3;
                const int endRow = 8;
                const int column = 1;
                const int total = endRow - startRow + 1;

                var sameStyle = true;
                var firstStyleId = ws.Cells[startRow, column].StyleID;
                var patternOk = 0;
                var colorOk = 0;

                for (var row = startRow; row <= endRow; row++)
                {
                    var cell = ws.Cells[row, column];
                    var styleId = cell.StyleID;
                    if (styleId != firstStyleId)
                    {
                        sameStyle = false;
                    }

                    var xfNode = stylesXml.SelectSingleNode($"//x:cellXfs/x:xf[{styleId + 1}]", ns);
                    var fillIdRaw = xfNode?.Attributes?["fillId"]?.Value;
                    if (!int.TryParse(fillIdRaw, out var fillId))
                    {
                        result.Errors.Add($"A{row}: khong doc duoc fillId tu style.");
                        continue;
                    }

                    var fillNode = stylesXml.SelectSingleNode($"//x:fills/x:fill[{fillId + 1}]", ns);
                    var patternNode = fillNode?.SelectSingleNode("x:patternFill", ns);
                    var patternType = patternNode?.Attributes?["patternType"]?.Value ?? string.Empty;
                    if (string.Equals(patternType, "lightTrellis", StringComparison.OrdinalIgnoreCase))
                    {
                        patternOk++;
                    }

                    var fgNode = patternNode?.SelectSingleNode("x:fgColor", ns);
                    var themeRaw = fgNode?.Attributes?["theme"]?.Value;
                    var tintRaw = fgNode?.Attributes?["tint"]?.Value;

                    var themeOk = string.Equals(themeRaw, "2", StringComparison.Ordinal);
                    var tintOk = true;
                    if (double.TryParse(tintRaw, NumberStyles.Any, CultureInfo.InvariantCulture, out var tint))
                    {
                        tintOk = Math.Abs(tint - (-0.1d)) <= 0.02d;
                    }

                    if (themeOk && tintOk)
                    {
                        colorOk++;
                    }
                }

                decimal score = 0;
                if (sameStyle)
                {
                    score += 1m;
                    result.Details.Add("A3:A8 co style dong nhat.");
                }
                else
                {
                    result.Errors.Add("A3:A8 khong dong nhat style, co dau hieu to mau khong deu.");
                }

                score += Math.Round(2m * patternOk / total, 2);
                score += Math.Round(1m * colorOk / total, 2);

                if (patternOk == total)
                {
                    result.Details.Add("A3:A8 da dung pattern fill lightTrellis.");
                }
                else
                {
                    result.Errors.Add($"Pattern fill chua dung cho toan bo A3:A8 ({patternOk}/{total}).");
                }

                if (colorOk != total)
                {
                    result.Errors.Add($"Mau fgColor chua khop theme/tint yeu cau ({colorOk}/{total}).");
                }
                else
                {
                    result.Details.Add("Mau pattern fill da khop theme/tint mong doi.");
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
