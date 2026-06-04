using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project11
{
    public class P11T3Grader : ITaskGrader
    {
        public string TaskId => "P11-T3";
        public string TaskName => "Doi ten sheet Outdoor Toys thanh Outdoor Sports";
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
                var workbook = studentSheet.Workbook;
                decimal score = 0m;

                var hasOutdoorSports = P11GraderHelpers.GetSheet(workbook, "Outdoor Sports") != null;
                var hasOutdoorToys = P11GraderHelpers.GetSheet(workbook, "Outdoor Toys") != null;

                if (hasOutdoorSports)
                {
                    score += 2m;
                    result.Details.Add("Da ton tai sheet moi 'Outdoor Sports'.");
                }
                else
                {
                    result.Errors.Add("Không tìm thấy sheet 'Outdoor Sports'.");
                }

                if (!hasOutdoorToys)
                {
                    score += 2m;
                    result.Details.Add("Không còn sheet cu 'Outdoor Toys'.");
                }
                else
                {
                    result.Errors.Add("Van con sheet cu 'Outdoor Toys'.");
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



