using System.Xml;
using System.Text.RegularExpressions;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project15
{
    public class P15T6Grader : ITaskGrader
    {
        public string TaskId => "P15-T6";
        public string TaskName => "Thiet lap mau tab worksheet la Pink Accent 1";
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
                var targetSheets = new[] { "Customers", "Products", "Orders" };
                var matched = 0;
                foreach (var name in targetSheets)
                {
                    var ws = P15GraderHelpers.GetSheet(studentSheet.Workbook, name);
                    if (ws == null)
                    {
                        continue;
                    }

                    var ns = new XmlNamespaceManager(ws.WorksheetXml.NameTable);
                    ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                    var tabColor = ws.WorksheetXml.SelectSingleNode("//x:sheetPr/x:tabColor", ns);
                    var theme = tabColor?.Attributes?["theme"]?.Value ?? string.Empty;
                    if (string.Equals(theme, "4", StringComparison.Ordinal))
                    {
                        matched++;
                    }
                }

                if (matched == targetSheets.Length)
                {
                    result.Score = MaxScore;
                    result.Details.Add("Da dat mau tab Pink Accent 1 cho Customers, Products, Orders.");
                }
                else
                {
                    result.Score = 2m;
                    result.Errors.Add($"Mau tab chưa đúng day du ({matched}/{targetSheets.Length} sheet).");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi: {ex.Message}");
            }

            return result;
        }
    }
}



