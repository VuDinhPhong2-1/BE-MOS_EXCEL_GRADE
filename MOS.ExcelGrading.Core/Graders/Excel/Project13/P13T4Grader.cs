using System.Globalization;
using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project13
{
    public class P13T4Grader : ITaskGrader
    {
        public string TaskId => "P13-T4";
        public string TaskName => "Shirt Orders: them dong Subtotal tai D201/F201";
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
                var ws = P13GraderHelpers.GetSheet(studentSheet.Workbook, "Shirt Orders");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Shirt Orders'.");
                    return result;
                }

                decimal score = 0m;
                var studentD201 = ws.Cells["D201"];
                if (P13GraderHelpers.CellMatchesExpected(studentD201, "191"))
                {
                    score += 2m;
                    result.Details.Add("D201 dung giá trị mong đợi (191).");
                }
                else
                {
                    result.Errors.Add(
                        $"D201 chưa đúng. Expected='191', Hiện tại='{studentD201.Text}'.");
                }

                var studentF201 = ws.Cells["F201"];
                var actualFormula = P13GraderHelpers.NormalizeFormula(studentF201.Formula);
                if (string.IsNullOrWhiteSpace(actualFormula)
                    && string.IsNullOrWhiteSpace((studentF201.Text ?? string.Empty).Trim()))
                {
                    score += 2m;
                    result.Details.Add("F201 dung yêu cầu de trong (khong co cong thuc, khong co giá trị).");
                }
                else
                {
                    result.Errors.Add(
                        $"F201 chưa đúng. Expected de trong, hien tai formula='{studentF201.Formula}', text='{studentF201.Text}'.");
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



