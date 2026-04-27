using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project22
{
    public class P22T2Grader : ITaskGrader
    {
        public string TaskId => "P22-T2";
        public string TaskName => "Trên trang tính \"Task\", đặt tên cho bảng là \"Task\".";
        public decimal MaxScore => 14m;

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
                var worksheet = P22GraderHelpers.GetSheet(studentSheet.Workbook, "Task");
                if (worksheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Task'.");
                    return result;
                }

                decimal score = 0m;

                if (worksheet.Tables.Count == 1)
                {
                    score += 2m;
                    result.Details.Add("Sheet Task chỉ có một bảng dữ liệu.");
                }
                else
                {
                    result.Errors.Add($"Sheet Task đang có {worksheet.Tables.Count} bảng dữ liệu, không đúng yêu cầu.");
                }

                var table = P22GraderHelpers.FindTable(
                    worksheet,
                    "Task",
                    "ID",
                    "Name",
                    "Task 1",
                    "Task 10",
                    "Total Tasks")
                    ?? worksheet.Tables.FirstOrDefault();

                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy bảng dữ liệu cần chấm trên sheet Task.");
                    return result;
                }

                var actualRange = P22GraderHelpers.NormalizeRange(table.Address.Address);
                if (P22GraderHelpers.IsRangeMatch(actualRange, "A3:M33"))
                {
                    score += 2m;
                    result.Details.Add("Phạm vi bảng dữ liệu vẫn đúng là A3:M33.");
                }
                else
                {
                    result.Errors.Add($"Phạm vi bảng dữ liệu chưa đúng. Hiện tại: {table.Address.Address}.");
                }

                var tableName = P22GraderHelpers.NormalizeText(table.Name);
                if (string.Equals(tableName, "Task", StringComparison.OrdinalIgnoreCase))
                {
                    score += 6m;
                    result.Details.Add("Tên bảng dữ liệu đã được đặt đúng là 'Task'.");
                }
                else
                {
                    result.Errors.Add($"Tên bảng dữ liệu chưa đúng. Hiện tại: '{tableName}', mong đợi: 'Task'.");
                }

                var displayName = P22GraderHelpers.NormalizeText(P22GraderHelpers.GetTableDisplayName(table));
                if (string.Equals(displayName, "Task", StringComparison.OrdinalIgnoreCase))
                {
                    score += 2m;
                    result.Details.Add("DisplayName của bảng dữ liệu đã đúng là 'Task'.");
                }
                else
                {
                    result.Errors.Add($"DisplayName của bảng dữ liệu chưa đúng. Hiện tại: '{displayName}', mong đợi: 'Task'.");
                }

                var hasRequiredStructure = table.Columns.Count == 13
                                           && table.Columns.Any(column =>
                                               string.Equals(
                                                   P22GraderHelpers.NormalizeIdentifier(column.Name),
                                                   "TOTALTASKS",
                                                   StringComparison.OrdinalIgnoreCase));
                if (hasRequiredStructure)
                {
                    score += 2m;
                    result.Details.Add("Cấu trúc cột của bảng vẫn đúng và không bị sai lệch.");
                }
                else
                {
                    result.Errors.Add("Cấu trúc cột của bảng đã bị thay đổi, chưa đúng với đề bài.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 2: {ex.Message}.");
            }

            return result;
        }
    }
}
