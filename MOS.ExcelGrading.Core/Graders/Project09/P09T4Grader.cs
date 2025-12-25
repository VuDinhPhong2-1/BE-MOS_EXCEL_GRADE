using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Project09
{
    public class P09T4Grader : ITaskGrader
    {
        public string TaskId => "P09-T4";
        public string TaskName => "Filter Total column: 34,000 to 45,000";
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
                decimal score = 0;

                // Rule 1: Có AutoFilter (2 điểm)
                if (studentSheet.AutoFilterAddress != null)
                {
                    score += 2m;
                    result.Details.Add($"✓ Đã bật AutoFilter tại {studentSheet.AutoFilterAddress}");

                    // Rule 2: Filter đúng cột Total với range 34000-45000 (2 điểm)
                    var filterColumn = GetFilterColumn(studentSheet, "Total");

                    if (filterColumn != null)
                    {
                        if (CheckFilterRange(studentSheet, filterColumn.Value, 34000, 45000))
                        {
                            score += 2m;
                            result.Details.Add("✓ Filter đúng khoảng 34,000 - 45,000");
                        }
                        else
                        {
                            result.Errors.Add("❌ Filter range không đúng hoặc có giá trị ngoài khoảng");
                        }
                    }
                    else
                    {
                        result.Errors.Add("❌ Không tìm thấy filter trên cột Total");
                    }
                }
                else
                {
                    result.Errors.Add("❌ Chưa bật AutoFilter");
                }

                result.Score = score;
            }
            catch (Exception ex)
            {
                result.Errors.Add($"❌ Lỗi: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Tìm cột có tên "Total" trong AutoFilter range
        /// </summary>
        private int? GetFilterColumn(ExcelWorksheet sheet, string columnName)
        {
            try
            {
                if (sheet.AutoFilterAddress == null)
                    return null;

                var range = sheet.Cells[sheet.AutoFilterAddress.Address];
                var headerRow = range.Start.Row;

                // Duyệt qua các cột trong AutoFilter range
                for (int col = range.Start.Column; col <= range.End.Column; col++)
                {
                    var cellValue = sheet.Cells[headerRow, col].Value?.ToString();

                    if (cellValue != null &&
                        cellValue.Contains(columnName, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"✅ Found '{columnName}' column at index {col}");
                        return col;
                    }
                }

                Console.WriteLine($"❌ Column '{columnName}' not found");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error finding column: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Kiểm tra các dòng visible có giá trị trong khoảng min-max
        /// </summary>
        private bool CheckFilterRange(ExcelWorksheet sheet, int column, double min, double max)
        {
            try
            {
                if (sheet.AutoFilterAddress == null)
                    return false;

                var autoFilterRange = sheet.Cells[sheet.AutoFilterAddress.Address];
                var startRow = autoFilterRange.Start.Row + 1; // Bỏ qua header
                var endRow = autoFilterRange.End.Row;

                int visibleCount = 0;
                int outOfRangeCount = 0;

                // Duyệt qua tất cả các dòng trong AutoFilter range
                for (int row = startRow; row <= endRow; row++)
                {
                    // Bỏ qua dòng bị ẩn
                    if (sheet.Row(row).Hidden)
                        continue;

                    visibleCount++;

                    var cell = sheet.Cells[row, column];
                    var cellValue = cell.Value;

                    // Thử parse giá trị
                    if (cellValue != null && double.TryParse(cellValue.ToString(), out double value))
                    {
                        // Kiểm tra có nằm trong khoảng không
                        if (value < min || value > max)
                        {
                            outOfRangeCount++;
                            Console.WriteLine($"⚠ Row {row}: Value {value} is out of range [{min}, {max}]");
                        }
                    }
                }

                Console.WriteLine($"Visible rows: {visibleCount}, Out of range: {outOfRangeCount}");

                // Nếu có dòng visible ngoài khoảng → Filter sai
                return outOfRangeCount == 0 && visibleCount > 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error checking filter range: {ex.Message}");
                return false;
            }
        }
    }
}
