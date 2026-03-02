using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Project01
{
    public class P01T5Grader : ITaskGrader
    {
        public string TaskId => "P01-T5";
        public string TaskName => "Tính Net Change (%) cho Units Sold và Gross Sales";
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

                var netChangeCells = GetNetChangeCells(studentMenu);
                if (netChangeCells.Count == 0)
                {
                    result.Errors.Add("Không xác định được vùng dữ liệu Net Change");
                    return result;
                }

                var filled = 0;
                var validPattern = 0;
                foreach (var addr in netChangeCells)
                {
                    var studentFormula = NormalizeFormula(studentMenu.Cells[addr].Formula);

                    if (!string.IsNullOrWhiteSpace(studentFormula))
                    {
                        filled++;
                        if (LooksLikeNetChangeFormula(studentFormula))
                        {
                            validPattern++;
                        }
                    }
                }

                var fillRatio = (decimal)filled / netChangeCells.Count;
                var validRatio = (decimal)validPattern / netChangeCells.Count;
                result.Score = Math.Round((fillRatio * 2m) + (validRatio * 2m), 2);

                result.Details.Add($"Net Change có công thức: {filled}/{netChangeCells.Count} ô");
                result.Details.Add($"Công thức có cấu trúc Net Change hợp lệ: {validPattern}/{netChangeCells.Count} ô");

                if (filled < netChangeCells.Count)
                {
                    result.Errors.Add("Một số ô Net Change chưa có công thức");
                }

                if (validPattern < netChangeCells.Count)
                {
                    result.Errors.Add("Một số công thức Net Change chưa đúng cấu trúc tăng/giảm phần trăm");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi: {ex.Message}");
            }

            return result;
        }

        private static List<string> GetNetChangeCells(ExcelWorksheet ws)
        {
            var cells = new List<string>();

            foreach (var table in ws.Tables)
            {
                var netChangeColumn = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name?.Trim(), "Net Change", StringComparison.OrdinalIgnoreCase));

                if (netChangeColumn == null)
                {
                    continue;
                }

                var colIndex = table.Address.Start.Column + netChangeColumn.Position;
                var startRow = table.Address.Start.Row + 1;
                var endRow = table.Address.End.Row;

                for (var row = startRow; row <= endRow; row++)
                {
                    cells.Add(ws.Cells[row, colIndex].Address);
                }
            }

            return cells;
        }

        private static bool LooksLikeNetChangeFormula(string normalizedFormula)
        {
            return normalizedFormula.Contains("/")
                && normalizedFormula.Contains("-")
                && normalizedFormula.Contains("(")
                && normalizedFormula.Contains(")");
        }

        private static string NormalizeFormula(string? formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return string.Empty;

            return formula
                .Replace("=", string.Empty)
                .Replace("$", string.Empty)
                .Replace(" ", string.Empty)
                .ToUpperInvariant();
        }
    }
}
