using OfficeOpenXml;
using OfficeOpenXml.Table;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Text.RegularExpressions;

namespace MOS.ExcelGrading.Core.Graders.Project01
{
    public class P01T5Grader : ITaskGrader
    {
        private static readonly HashSet<string> TargetTableAddresses =
            new(StringComparer.OrdinalIgnoreCase)
            {
                "B5:F10",   // Units_Sold
                "B13:F18"   // Gross Sales (Table3)
            };

        public string TaskId => "P01-T5";
        public string TaskName => "Tính % Change cho Units Sold và Gross Sales";
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

                var targetTables = studentMenu.Tables
                    .Where(t => TargetTableAddresses.Contains(t.Address.Address))
                    .ToList();

                if (targetTables.Count == 0)
                {
                    result.Errors.Add("Không tìm thấy 2 bảng mục tiêu để chấm % Change (B5:F10, B13:F18)");
                    return result;
                }

                var percentCells = GetPercentChangeCells(studentMenu, targetTables);
                if (percentCells.Count == 0)
                {
                    result.Errors.Add("Không xác định được cột '% Change' trong bảng mục tiêu");
                    return result;
                }

                var filled = 0;
                var validPattern = 0;

                foreach (var addr in percentCells)
                {
                    var formula = NormalizeFormula(studentMenu.Cells[addr].Formula);
                    if (string.IsNullOrWhiteSpace(formula))
                    {
                        continue;
                    }

                    filled++;
                    if (LooksLikePercentChangeFormula(formula))
                    {
                        validPattern++;
                    }
                }

                var tableCoverageScore = Math.Min(1m, (decimal)targetTables.Count / TargetTableAddresses.Count);
                var fillScore = Math.Round(((decimal)filled / percentCells.Count) * 1.5m, 2);
                var validScore = Math.Round(((decimal)validPattern / percentCells.Count) * 1.5m, 2);

                result.Score = Math.Min(MaxScore, tableCoverageScore + fillScore + validScore);

                result.Details.Add($"Bảng mục tiêu tìm thấy: {targetTables.Count}/{TargetTableAddresses.Count}");
                result.Details.Add($"% Change có công thức: {filled}/{percentCells.Count} ô");
                result.Details.Add($"Công thức % Change đúng cấu trúc: {validPattern}/{percentCells.Count} ô");

                if (targetTables.Count < TargetTableAddresses.Count)
                {
                    result.Errors.Add("Thiếu một trong hai bảng mục tiêu cần tính % Change");
                }

                if (filled < percentCells.Count)
                {
                    result.Errors.Add("Một số ô % Change chưa có công thức");
                }

                if (validPattern < percentCells.Count)
                {
                    result.Errors.Add("Một số công thức % Change chưa đúng dạng tăng/giảm theo phần trăm");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi: {ex.Message}");
            }

            return result;
        }

        private static List<string> GetPercentChangeCells(
            ExcelWorksheet worksheet,
            IEnumerable<ExcelTable> tables)
        {
            var cells = new List<string>();

            foreach (var table in tables)
            {
                var percentCol = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name?.Trim(), "% Change", StringComparison.OrdinalIgnoreCase));
                var netChangeCol = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name?.Trim(), "Net Change", StringComparison.OrdinalIgnoreCase));
                var year2016Col = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name?.Trim(), "2016", StringComparison.OrdinalIgnoreCase));

                if (percentCol == null || netChangeCol == null || year2016Col == null)
                {
                    continue;
                }

                var colIndex = table.Address.Start.Column + percentCol.Position;
                var startRow = table.Address.Start.Row + 1;
                var endRow = table.Address.End.Row;

                for (var row = startRow; row <= endRow; row++)
                {
                    cells.Add(worksheet.Cells[row, colIndex].Address);
                }
            }

            return cells;
        }

        private static bool LooksLikePercentChangeFormula(string normalizedFormula)
        {
            var usesStructuredRefs = normalizedFormula.Contains("/")
                && normalizedFormula.Contains("NETCHANGE")
                && normalizedFormula.Contains("2016");

            if (usesStructuredRefs)
            {
                return true;
            }

            // Chấp nhận công thức ô-thường dạng E6/C6 nếu học viên nhập theo tham chiếu trực tiếp.
            return Regex.IsMatch(
                normalizedFormula,
                @"^\(?[A-Z]{1,3}\d+\)?/\(?[A-Z]{1,3}\d+\)?$",
                RegexOptions.CultureInvariant);
        }

        private static string NormalizeFormula(string? formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                return string.Empty;
            }

            return formula
                .Replace("=", string.Empty)
                .Replace("$", string.Empty)
                .Replace(" ", string.Empty)
                .ToUpperInvariant();
        }
    }
}
