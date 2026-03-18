using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using System.Xml;

namespace MOS.ExcelGrading.Core.Graders.Project11
{
    public class P11T5Grader : ITaskGrader
    {
        public string TaskId => "P11-T5";
        public string TaskName => "Costs: set print titles rows 1:3";
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

                if (printTitleNode != null)
                {
                    score += 1m;
                    result.Details.Add("Tim thay defined name _xlnm.Print_Titles.");
                }
                else
                {
                    result.Errors.Add("Khong tim thay defined name _xlnm.Print_Titles.");
                    result.Score = score;
                    return result;
                }

                var localSheetIdText = printTitleNode.Attributes?["localSheetId"]?.Value ?? string.Empty;
                var costsSheetIndex = P11GraderHelpers.GetSheetIndex0Based(workbook, "Costs");
                if (int.TryParse(localSheetIdText, out var localSheetId)
                    && costsSheetIndex >= 0
                    && localSheetId == costsSheetIndex)
                {
                    score += 1m;
                    result.Details.Add("localSheetId cua Print_Titles dung voi sheet Costs.");
                }
                else
                {
                    result.Errors.Add(
                        $"localSheetId chua dung. Hien tai: '{localSheetIdText}', mong doi: {costsSheetIndex}.");
                }

                var definedValue = printTitleNode.InnerText ?? string.Empty;
                var normalizedDefinedValue = (definedValue ?? string.Empty)
                    .Replace("$", string.Empty, StringComparison.Ordinal)
                    .Replace("'", string.Empty, StringComparison.Ordinal)
                    .Replace(" ", string.Empty, StringComparison.Ordinal)
                    .ToUpperInvariant();

                if (string.Equals(normalizedDefinedValue, "COSTS!1:3", StringComparison.Ordinal))
                {
                    score += 2m;
                    result.Details.Add("Print titles dung: Costs!$1:$3.");
                }
                else
                {
                    result.Errors.Add($"Gia tri Print_Titles chua dung. Hien tai: '{definedValue}'.");
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
