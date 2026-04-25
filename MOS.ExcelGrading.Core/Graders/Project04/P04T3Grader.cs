using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Sparkline;

namespace MOS.ExcelGrading.Core.Graders.Project04
{
    public class P04T3Grader : ITaskGrader
    {
        public string TaskId => "P04-T3";
        public string TaskName => "Chèn Sparkline Line tai G5:G25 tren Enrollment";
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
                var ws = P04GraderHelpers.GetSheet(studentSheet, "Enrollment");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet Enrollment");
                    return result;
                }

                if (ws.SparklineGroups.Count == 0)
                {
                    result.Errors.Add("Không tìm thấy sparkline group");
                    return result;
                }

                result.Score += 1m;
                result.Details.Add($"Tìm thấy {ws.SparklineGroups.Count} sparkline group");

                var targetGroup = ws.SparklineGroups.FirstOrDefault(g =>
                    P04GraderHelpers.NormalizeAddress(g.LocationRange?.Address) == "G5:G25");

                if (targetGroup == null)
                {
                    result.Errors.Add("Không tìm thấy sparkline tại vùng G5:G25");
                    result.Score = Math.Min(MaxScore, result.Score);
                    return result;
                }

                if (targetGroup.Type == eSparklineType.Line && targetGroup.Sparklines.Count == 21)
                {
                    result.Score += 1.5m;
                    result.Details.Add("Sparkline đúng loại Line và đủ số dòng (21)");
                }
                else
                {
                    result.Errors.Add($"Sparkline chưa đúng loại/ số dòng (type={targetGroup.Type}, count={targetGroup.Sparklines.Count})");
                }

                var dataAddress = P04GraderHelpers.NormalizeAddress(targetGroup.DataRange?.Address);
                if (dataAddress == "D5:F25")
                {
                    result.Score += 1.5m;
                    result.Details.Add("Nguồn dữ liệu sparkline đúng D5:F25");
                }
                else
                {
                    result.Errors.Add($"Nguồn dữ liệu sparkline chưa đúng. Hiện tại: {targetGroup.DataRange?.Address}");
                }

                result.Score = Math.Min(MaxScore, result.Score);
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
