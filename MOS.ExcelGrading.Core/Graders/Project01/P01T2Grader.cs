using OfficeOpenXml;
using OfficeOpenXml.Table;
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
                const string targetAddress = "B5:F10";
                const string targetName = "Units_Sold";
                const string oldName = "Table2";

                var studentMenu = studentSheet.Workbook.Worksheets["Menu Items"];
                if (studentMenu == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Menu Items'");
                    return result;
                }

                decimal score = 0;
                var studentTables = studentMenu.Tables.ToList();
                var targetTable = FindTableByAddress(studentTables, targetAddress);

                // 1) Xác định đúng bảng đích theo vùng dữ liệu gốc của đề.
                if (targetTable != null)
                {
                    score += 1;
                    result.Details.Add($"Tìm thấy bảng mục tiêu tại vùng {targetAddress}");
                }
                else
                {
                    result.Errors.Add($"Không tìm thấy bảng mục tiêu tại vùng {targetAddress}");
                }

                // 2) Bảng ở vùng đích phải được đổi tên thành Units_Sold.
                if (targetTable != null &&
                    targetTable.Name.Equals(targetName, StringComparison.Ordinal))
                {
                    score += 2;
                    result.Details.Add($"Bảng tại {targetAddress} đã đổi tên đúng thành '{targetName}'");
                }
                else
                {
                    var currentName = targetTable?.Name ?? "(không xác định)";
                    result.Errors.Add($"Bảng tại {targetAddress} chưa đổi đúng tên. Hiện tại: '{currentName}'");
                }

                // 3) Không còn tên cũ Table2 trong Menu Items.
                var hasOldName = studentTables.Any(t =>
                    t.Name.Equals(oldName, StringComparison.Ordinal));
                if (!hasOldName)
                {
                    score += 1;
                    result.Details.Add($"Không còn bảng tên cũ '{oldName}'");
                }
                else
                {
                    result.Errors.Add($"Vẫn còn bảng tên cũ '{oldName}'");
                }

                // 4) Anti-cheat: Units_Sold phải thuộc đúng bảng đích, không phải bảng khác.
                var unitsSoldTables = studentTables
                    .Where(t => t.Name.Equals(targetName, StringComparison.Ordinal))
                    .ToList();

                if (unitsSoldTables.Count == 1 &&
                    targetTable != null &&
                    ReferenceEquals(unitsSoldTables[0], targetTable))
                {
                    result.Details.Add($"Tên '{targetName}' được gán đúng cho bảng mục tiêu");
                }
                else
                {
                    result.Errors.Add(
                        $"Tên '{targetName}' đang gán sai bảng hoặc bị trùng (số bảng trùng tên: {unitsSoldTables.Count})");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi: {ex.Message}");
            }

            return result;
        }

        private static ExcelTable? FindTableByAddress(IEnumerable<ExcelTable> tables, string expectedAddress)
        {
            return tables.FirstOrDefault(t =>
                t.Address.Address.Equals(expectedAddress, StringComparison.OrdinalIgnoreCase));
        }
    }
}
