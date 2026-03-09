using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project07
{
    public class P07T2Grader : ITaskGrader
    {
        public string TaskId => "P07-T2";
        public string TaskName => "Tea table style = Blue, Table Style Medium 9";
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
                var ws = P07GraderHelpers.GetSheet(studentSheet, "Tea");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Tea'.");
                    return result;
                }

                var table = ws.Tables.FirstOrDefault();
                if (table == null)
                {
                    result.Errors.Add("Khong tim thay table tren sheet Tea.");
                    return result;
                }

                decimal score = 1m; // Tim thay table.
                var styleOk = table.TableStyle == TableStyles.Medium9;
                if (styleOk)
                {
                    score += 2m;
                    result.Details.Add("Table style dung Medium 9.");
                }
                else
                {
                    result.Errors.Add($"Table style chua dung. Hien tai: {table.TableStyle}.");
                }

                var addressOk = P07GraderHelpers.NormalizeAddress(table.Address.Address) == "A7:D15";
                if (addressOk)
                {
                    score += 1m;
                    result.Details.Add("Table address dung A7:D15.");
                }
                else
                {
                    result.Errors.Add($"Table address khong dung. Hien tai: {table.Address.Address}.");
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
