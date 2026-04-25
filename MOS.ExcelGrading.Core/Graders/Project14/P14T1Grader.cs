using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project14
{
    public class P14T1Grader : ITaskGrader
    {
        public string TaskId => "P14-T1";
        public string TaskName => "January: thiet lap Print Area A4:F20";
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

                var januaryIndex = workbook.Worksheets.ToList()
                    .FindIndex(ws => string.Equals(ws.Name, "January", StringComparison.OrdinalIgnoreCase));
                if (januaryIndex < 0)
                {
                    result.Errors.Add("Không tìm thấy sheet 'January'.");
                    return result;
                }

                decimal score = 0m;
                var ns = new XmlNamespaceManager(workbookXml.NameTable);
                ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var printAreaNode = workbookXml.SelectSingleNode(
                    $"//x:definedNames/x:definedName[@name='_xlnm.Print_Area' and @localSheetId='{januaryIndex}']",
                    ns);

                if (printAreaNode != null)
                {
                    score += 2m;
                    result.Details.Add("Tìm thấy Print_Area cho sheet January.");
                }
                else
                {
                    result.Errors.Add("Không tìm thấy Print_Area cho sheet January.");
                    result.Score = score;
                    return result;
                }

                var normalized = P14GraderHelpers.NormalizePrintArea(printAreaNode.InnerText);
                var isExpected = string.Equals(normalized, "JANUARY!A4:F20", StringComparison.Ordinal)
                                 || string.Equals(normalized, "TABLE4[[#ALL],[CLIENTID]:[POLICYTYPE]]", StringComparison.Ordinal);
                if (isExpected)
                {
                    score += 2m;
                    result.Details.Add("Print_Area dung phạm vi yêu cầu A4:F20.");
                }
                else
                {
                    result.Errors.Add($"Gia tri Print_Area chưa đúng. Hiện tại: '{printAreaNode.InnerText}'.");
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



