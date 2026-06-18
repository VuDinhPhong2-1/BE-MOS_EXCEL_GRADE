using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project22
{
    public class P22T1Grader : ITaskGrader
    {
        public string TaskId => "P22-T1";
        public string TaskName => "Sao chép định dạng từ tiêu đề và phụ đề của trang tính \"Task\", sau đó áp dụng định dạng đó cho tiêu đề và phụ đề của trang tính \"Project\".";
        public decimal MaxScore => 18m;

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
                var taskSheet = P22GraderHelpers.GetSheet(studentSheet.Workbook, "Task");
                var projectSheet = P22GraderHelpers.GetSheet(studentSheet.Workbook, "Project");
                if (taskSheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Task'.");
                    return result;
                }

                if (projectSheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Project'.");
                    return result;
                }

                decimal score = 0m;

                var taskTitleText = P22GraderHelpers.NormalizeText(taskSheet.Cells["A1"].Text);
                var taskSubtitleText = P22GraderHelpers.NormalizeText(taskSheet.Cells["A2"].Text);
                if (!string.IsNullOrWhiteSpace(taskTitleText) && !string.IsNullOrWhiteSpace(taskSubtitleText))
                {
                    score += 2m;
                    result.Details.Add("Đã xác nhận được tiêu đề và phụ đề nguồn tại Task!A1:A2.");
                }
                else
                {
                    result.Errors.Add("Không đọc được đầy đủ nội dung nguồn tại Task!A1:A2.");
                }

                var projectTitleText = P22GraderHelpers.NormalizeText(projectSheet.Cells["A1"].Text);
                var projectSubtitleText = P22GraderHelpers.NormalizeText(projectSheet.Cells["A2"].Text);
                if (string.Equals(projectTitleText, taskTitleText, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(projectSubtitleText, taskSubtitleText, StringComparison.OrdinalIgnoreCase))
                {
                    score += 3m;
                    result.Details.Add("Nội dung tiêu đề và phụ đề trên sheet Project vẫn được giữ đúng.");
                }
                else
                {
                    result.Errors.Add(
                        $"Nội dung tiêu đề/phụ đề trên Project chưa đúng. A1='{projectTitleText}', A2='{projectSubtitleText}'.");
                }

                var sourceTitleStyle = taskSheet.Cells["A1"].StyleID;
                var sourceSubtitleStyle = taskSheet.Cells["A2"].StyleID;
                var targetTitleStyle = projectSheet.Cells["A1"].StyleID;
                var targetSubtitleStyle = projectSheet.Cells["A2"].StyleID;

                if (targetTitleStyle == sourceTitleStyle)
                {
                    score += 4m;
                    result.Details.Add("Project!A1 đã nhận đúng định dạng từ Task!A1.");
                }
                else
                {
                    result.Errors.Add(
                        $"Project!A1 chưa đúng định dạng. Hiện tại style={targetTitleStyle}, mong đợi style={sourceTitleStyle}.");
                }

                if (targetSubtitleStyle == sourceSubtitleStyle)
                {
                    score += 4m;
                    result.Details.Add("Project!A2 đã nhận đúng định dạng từ Task!A2.");
                }
                else
                {
                    result.Errors.Add(
                        $"Project!A2 chưa đúng định dạng. Hiện tại style={targetSubtitleStyle}, mong đợi style={sourceSubtitleStyle}.");
                }

                if (targetTitleStyle != targetSubtitleStyle)
                {
                    score += 2m;
                    result.Details.Add("Định dạng tiêu đề và phụ đề trên sheet Project đã được tách đúng thành hai kiểu khác nhau.");
                }
                else
                {
                    result.Errors.Add("Project!A1 và Project!A2 đang cùng một kiểu định dạng, chưa phân biệt title và subtitle.");
                }

                var projectA3Style = projectSheet.Cells["A3"].StyleID;
                var projectB3Style = projectSheet.Cells["B3"].StyleID;
                if (projectA3Style == projectB3Style
                    && projectA3Style != targetTitleStyle
                    && projectA3Style != targetSubtitleStyle)
                {
                    score += 3m;
                    result.Details.Add("Vùng dữ liệu tiếp theo (hàng 3) không bị áp nhầm định dạng tiêu đề/phụ đề.");
                }
                else
                {
                    result.Errors.Add(
                        $"Có dấu hiệu áp định dạng sai phạm vi quanh hàng 3. Style A3={projectA3Style}, B3={projectB3Style}.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 1: {ex.Message}.");
            }

            return result;
        }
    }
}

