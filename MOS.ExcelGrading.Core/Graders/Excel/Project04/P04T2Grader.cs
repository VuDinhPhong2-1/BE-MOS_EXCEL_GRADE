using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project04
{
    public class P04T2Grader : ITaskGrader
    {
        public string TaskId => "P04-T2";
        public string TaskName => "Đặt độ rộng cột B:G = 12 trên sheet Number of course hours";
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
                    result.Errors.Add("Không tìm thấy sheet Number of course hours");
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
                        result.Errors.Add($"Cột {c} có độ rộng {width:0.###}, chưa đặt giá trị 12");
                    }
                }

                result.Score = Math.Round(MaxScore * okCount / 6m, 2);
                if (okCount == 6)
                {
                    result.Details.Add("Tất cả cột B:G đã đặt độ rộng 12");
                }
                else
                {
                    result.Details.Add($"Cột đặt đúng: {okCount}/6");
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


// minor-sync: non-functional graders update
