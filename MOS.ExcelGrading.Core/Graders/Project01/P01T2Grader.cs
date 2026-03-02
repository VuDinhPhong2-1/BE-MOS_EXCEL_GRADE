using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Project01
{
    public class P01T2Grader : ITaskGrader
    {
        public string TaskId => "P01-T2";
        public string TaskName => "Đổi tên bảng Table2 thành Units_Sold trong Menu Items";
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
                var studentMenu = studentSheet.Workbook.Worksheets["Menu Items"];

                if (studentMenu == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Menu Items'");
                    return result;
                }

                decimal score = 0;
                var studentTables = studentMenu.Tables.ToList();
                var studentNames = studentTables
                    .Select(t => t.Name)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                // Điều kiện 1: phải có Units_Sold
                if (studentNames.Contains("Units_Sold"))
                {
                    score += 1;
                    result.Details.Add("Có bảng tên 'Units_Sold'");
                }
                else
                {
                    result.Errors.Add("Không tìm thấy bảng tên 'Units_Sold'");
                }

                // Điều kiện 2: không còn Table2
                if (!studentNames.Contains("Table2"))
                {
                    score += 1;
                    result.Details.Add("Không còn bảng tên 'Table2'");
                }
                else
                {
                    result.Errors.Add("Vẫn còn bảng tên 'Table2'");
                }

                // Điều kiện 3: các bảng còn lại vẫn giữ Table1/Table3/Table4
                var expectedOtherNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "Table1",
                    "Table3",
                    "Table4"
                };

                if (expectedOtherNames.All(studentNames.Contains))
                {
                    score += 1;
                    result.Details.Add("Các bảng còn lại giữ đúng tên: Table1, Table3, Table4");
                }
                else
                {
                    result.Errors.Add("Tên các bảng còn lại không đúng bộ Table1/Table3/Table4");
                }

                // Điều kiện 4: số lượng bảng giữ nguyên (chỉ đổi tên, không thêm/bớt bảng)
                var studentUnitsSold = studentMenu.Tables.FirstOrDefault(t =>
                    t.Name.Equals("Units_Sold", StringComparison.OrdinalIgnoreCase));

                if (studentUnitsSold != null && studentTables.Count == 4)
                {
                    score += 1;
                    result.Details.Add("Số lượng bảng giữ nguyên sau khi đổi tên");
                }
                else
                {
                    result.Errors.Add("Số lượng bảng thay đổi hoặc chưa đổi tên đúng");
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
