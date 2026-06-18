using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project08
{
    public class P08T2Grader : ITaskGrader
    {
        public string TaskId => "P08-T2";
        public string TaskName => "Sale History bật chế độ Show Formulas";
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
                var ws = P08GraderHelpers.GetSheet(studentSheet, "Sale History");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Sale History'.");
                    return result;
                }

                decimal score = 1m;
                var ns = new XmlNamespaceManager(ws.WorksheetXml.NameTable);
                ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var viewNode = ws.WorksheetXml.SelectSingleNode("//x:sheetViews/x:sheetView", ns);
                if (viewNode == null)
                {
                    result.Errors.Add("Không đọc được sheetView để kiểm tra Show Formulas.");
                    result.Score = score;
                    return result;
                }

                score += 1m;
                var raw = viewNode.Attributes?["showFormulas"]?.Value ?? string.Empty;
                var enabled = string.Equals(raw, "1", StringComparison.OrdinalIgnoreCase) ||
                              string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase);
                if (enabled)
                {
                    score += 2m;
                    result.Details.Add("Sheet Sale History đã bật Show Formulas.");
                }
                else
                {
                    result.Errors.Add("Sheet Sale History chưa bật Show Formulas.");
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

