using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project04
{
    public class P04T2Grader : ITaskGrader
    {
        public string TaskId => "P04-T2";
        public string TaskName => "Dat do rong cot B:G = 12 tren Number of course hours";
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
                var ws = P04GraderHelpers.GetSheet(studentSheet, "Number of course hours");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet Number of course hours");
                    return result;
                }

                var okCount = 0;
                for (var c = 2; c <= 7; c++)
                {
                    var width = ws.Column(c).Width;
                    if (P04GraderHelpers.IsWidth12(width))
                    {
                        okCount++;
                    }
                    else
                    {
                        result.Errors.Add($"Cot {c} co do rong {width:0.###}, chua dat gia tri 12");
                    }
                }

                result.Score = Math.Round(MaxScore * okCount / 6m, 2);
                if (okCount == 6)
                {
                    result.Details.Add("Tat ca cot B:G da dat do rong 12");
                }
                else
                {
                    result.Details.Add($"Cot dat dung: {okCount}/6");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }
}

