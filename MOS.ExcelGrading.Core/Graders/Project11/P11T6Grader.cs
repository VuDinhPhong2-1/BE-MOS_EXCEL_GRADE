using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project11
{
    public class P11T6Grader : ITaskGrader
    {
        public string TaskId => "P11-T6";
        public string TaskName => "Costs: thiet lap Print Titles hang 1:3";
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
                    result.Errors.Add("Không đọc được workbook XML.");
                    return result;
                }

                decimal score = 0m;
                var ns = P11GraderHelpers.CreateWorkbookNamespaceManager(workbookXml);
                var printTitleNode = workbookXml.SelectSingleNode(
                    "//x:definedNames/x:definedName[@name='_xlnm.Print_Titles']",
                    ns);
                if (printTitleNode == null)
                {
                    result.Errors.Add("Không tìm thấy defined name _xlnm.Print_Titles.");
                    return result;
                }

                score += 1m;
                result.Details.Add("Tìm thấy defined name _xlnm.Print_Titles.");

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
                    result.Errors.Add($"localSheetId chưa đúng. Hiện tại: '{localSheetIdText}', mong đợi: {costsSheetIndex}.");
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
                    result.Errors.Add($"Gia tri Print Titles chưa đúng. Hiện tại: '{printTitleNode.InnerText}'.");
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



