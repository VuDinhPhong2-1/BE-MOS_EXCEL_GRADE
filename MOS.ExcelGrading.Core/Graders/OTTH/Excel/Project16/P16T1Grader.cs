using System.Xml;
using System.Text.RegularExpressions;
using System.Globalization;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project16
{
    public class P16T1Grader : ITaskGrader
    {
        public string TaskId => "P16-T1";
        public string TaskName => "Products: co dinh 2 hang dau";
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
                var ws = P16GraderHelpers.GetSheet(studentSheet.Workbook, "Products");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'Products'.");
                    return result;
                }

                decimal score = 0m;
                var ns = new XmlNamespaceManager(ws.WorksheetXml.NameTable);
                ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var pane = ws.WorksheetXml.SelectSingleNode("//x:sheetViews/x:sheetView/x:pane", ns);
                var ySplit = pane?.Attributes?["ySplit"]?.Value ?? string.Empty;
                var topLeft = pane?.Attributes?["topLeftCell"]?.Value ?? string.Empty;
                var state = pane?.Attributes?["state"]?.Value ?? string.Empty;

                if (string.Equals(ySplit, "2", StringComparison.Ordinal))
                {
                    score += 2m;
                    result.Details.Add("ySplit dung = 2 (dòng 1-2 duoc co dinh).");
                }
                else
                {
                    result.Errors.Add($"ySplit chua dúng. Hi?n t?i: '{ySplit}'.");
                }

                if (string.Equals(topLeft, "A3", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(state, "frozen", StringComparison.OrdinalIgnoreCase))
                {
                    score += 2m;
                    result.Details.Add("TopLeftCell va state dung (A3, frozen).");
                }
                else
                {
                    result.Errors.Add($"Trang thai pane chua dúng. TopLeftCell='{topLeft}', State='{state}'.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"L?i: {ex.Message}");
            }

            return result;
        }
    }
}




