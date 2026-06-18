using System.Globalization;
using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project13
{
    public class P13T2Grader : ITaskGrader
    {
        public string TaskId => "P13-T2";
        public string TaskName => "Shirt Orders: cong thuc C2 SUMIF cho Blue cost";
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
                var ws = P13GraderHelpers.GetSheet(studentSheet.Workbook, "Shirt Orders");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'Shirt Orders'.");
                    return result;
                }

                var isMatched = P13GraderHelpers.FormulaMatchesAny(
                    ws.Cells["C2"].Formula,
                    "SUMIF(D:D,\"Blue\",F:F)",
                    "SUMIF(D6:D199,\"Blue\",F6:F199)");

                if (isMatched)
                {
                    result.Score = MaxScore;
                    result.Details.Add("Công th?c C2 dung chinh xac.");
                }
                else
                {
                    result.Errors.Add(
                        $"Công th?c C2 chua dúng. Chap nhan SUMIF(D:D,\"Blue\",F:F) hoac SUMIF(D6:D199,\"Blue\",F6:F199). Hi?n t?i: '{ws.Cells["C2"].Formula}'.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"L?i: {ex.Message}");
            }

            return result;
        }
    }
}




