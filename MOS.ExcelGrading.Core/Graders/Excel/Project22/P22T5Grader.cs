using System.Text.RegularExpressions;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project22
{
    public class P22T5Grader : ITaskGrader
    {
        public string TaskId => "P22-T5";
        public string TaskName => "Trên trang tính \"Exams\", tại ô E35, sử dụng công thức để xác định số học sinh không đạt điểm Exam 3.";
        public decimal MaxScore => 22m;

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
                var worksheet = P22GraderHelpers.GetSheet(studentSheet.Workbook, "Exams");
                if (worksheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Exams'.");
                    return result;
                }

                decimal score = 0m;

                var table = P22GraderHelpers.FindTable(
                    worksheet,
                    "Table3",
                    "ID",
                    "Name",
                    "Exam 1",
                    "Exam 2",
                    "Exam 3",
                    "Total Exams")
                    ?? P22GraderHelpers.FindTable(
                        worksheet,
                        null,
                        "ID",
                        "Name",
                        "Exam 1",
                        "Exam 2",
                        "Exam 3",
                        "Total Exams");

                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy bảng dữ liệu Exam để chấm Task 5.");
                    return result;
                }

                if (!P22GraderHelpers.TryGetColumnIndex(table, "Exam 3", out var exam3Column))
                {
                    result.Errors.Add("Không xác định được cột 'Exam 3' trong bảng dữ liệu.");
                    return result;
                }

                score += 3m;
                result.Details.Add("Đã tìm thấy đúng cột 'Exam 3' trong bảng dữ liệu.");

                var targetCell = worksheet.Cells["E35"];
                var formula = targetCell.Formula ?? string.Empty;
                var normalizedFormula = P22GraderHelpers.NormalizeFormula(formula);

                if (!string.IsNullOrWhiteSpace(formula))
                {
                    score += 4m;
                    result.Details.Add("Ô E35 đã có công thức.");
                }
                else
                {
                    result.Errors.Add("Ô E35 chưa có công thức.");
                    result.Score = score;
                    return result;
                }

                if (normalizedFormula.Contains("COUNTBLANK(", StringComparison.Ordinal))
                {
                    score += 5m;
                    result.Details.Add("Công thức tại E35 đã sử dụng hàm COUNTBLANK.");
                }
                else
                {
                    result.Errors.Add($"Công thức tại E35 chưa dùng hàm COUNTBLANK. Hiện tại: '{formula}'.");
                }

                var usesStructuredExam3Reference =
                    normalizedFormula.Contains("[EXAM3]", StringComparison.Ordinal)
                    || Regex.IsMatch(
                        normalizedFormula,
                        @"COUNTBLANK\([^)]*\[[^]]*EXAM3[^]]*][^)]*\)",
                        RegexOptions.CultureInvariant);

                if (usesStructuredExam3Reference)
                {
                    score += 5m;
                    result.Details.Add("Công thức đã tham chiếu đúng cột Exam 3 bằng cấu trúc cột của bảng.");
                }
                else
                {
                    result.Errors.Add("Công thức chưa tham chiếu đúng cột Exam 3 theo dạng cấu trúc cột của bảng.");
                }

                var hasDirectRangeReference =
                    Regex.IsMatch(normalizedFormula, @"[A-Z]{1,3}\d+:[A-Z]{1,3}\d+", RegexOptions.CultureInvariant)
                    || Regex.IsMatch(normalizedFormula, @"[A-Z]{1,3}:[A-Z]{1,3}", RegexOptions.CultureInvariant);

                if (!hasDirectRangeReference)
                {
                    score += 3m;
                    result.Details.Add("Công thức không dùng tham chiếu phạm vi ô trực tiếp.");
                }
                else
                {
                    result.Errors.Add("Công thức đang dùng tham chiếu phạm vi ô trực tiếp thay vì cấu trúc cột.");
                }

                var firstDataRow = table.Address.Start.Row + 1;
                var lastDataRow = table.Address.End.Row;
                var blankCount = 0;
                for (var row = firstDataRow; row <= lastDataRow; row++)
                {
                    var cell = worksheet.Cells[row, exam3Column];
                    if (string.IsNullOrWhiteSpace(cell.Text))
                    {
                        blankCount++;
                    }
                }

                if (P22GraderHelpers.TryGetNumericValue(targetCell, out var actualBlankCount)
                    && actualBlankCount == blankCount)
                {
                    score += 5m;
                    result.Details.Add("Giá trị kết quả tại E35 khớp đúng với số ô trống của cột Exam 3.");
                }
                else
                {
                    result.Errors.Add(
                        $"Giá trị kết quả tại E35 chưa đúng. Hiện tại: '{targetCell.Text}', mong đợi: {blankCount}.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 5: {ex.Message}.");
            }

            return result;
        }
    }
}
