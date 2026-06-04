using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project08
{
    public class P08T5Grader : ITaskGrader
    {
        public string TaskId => "P08-T5";
        public string TaskName => "Sales Postal Code đúng UPPER cho 3 ký tự đầu";
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
                var ws = P08GraderHelpers.GetSheet(studentSheet, "Sales");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Sales'.");
                    return result;
                }

                var table = ws.Tables.FirstOrDefault(t =>
                    t.Columns.Any(c => (c.Name ?? string.Empty).Contains("Postal", StringComparison.OrdinalIgnoreCase)));
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy table chứa cột Postal Code.");
                    return result;
                }

                var postalCol = table.Columns.First(c =>
                    (c.Name ?? string.Empty).Contains("Postal", StringComparison.OrdinalIgnoreCase));
                var addressCol = table.Columns.FirstOrDefault(c =>
                    string.Equals((c.Name ?? string.Empty).Trim(), "Address", StringComparison.OrdinalIgnoreCase));

                var startRow = table.Address.Start.Row + 1;
                var endRow = table.Address.End.Row - (table.ShowTotal ? 1 : 0);
                var rowCount = Math.Max(0, endRow - startRow + 1);
                if (rowCount == 0)
                {
                    result.Errors.Add("Không có dữ liệu để chấm.");
                    return result;
                }

                var postalColIndex = table.Address.Start.Column + postalCol.Position;
                var addressColIndex = addressCol != null
                    ? table.Address.Start.Column + addressCol.Position
                    : 4;

                var formulaRows = 0;
                var logicRows = 0;
                var valueRows = 0;

                for (var row = startRow; row <= endRow; row++)
                {
                    var postalCell = ws.Cells[row, postalColIndex];
                    var formulaRaw = postalCell.Formula ?? string.Empty;
                    var formula = P08GraderHelpers.NormalizeFormula(formulaRaw);

                    if (!string.IsNullOrWhiteSpace(formulaRaw))
                    {
                        formulaRows++;
                    }
                    else
                    {
                        result.Errors.Add($"Hang {row}: Postal Code chưa có công thức.");
                    }

                    var addressRef = OfficeOpenXml.ExcelCellBase.GetAddress(row, addressColIndex);
                    var hasUpper = formula.Contains("UPPER(", StringComparison.Ordinal);
                    var hasLeft = formula.Contains("LEFT(", StringComparison.Ordinal);
                    var hasLength3 = formula.Contains(",3", StringComparison.Ordinal);
                    var hasAddressRef = formula.Contains("ADDRESS", StringComparison.Ordinal) ||
                                        formula.Contains(addressRef, StringComparison.OrdinalIgnoreCase);
                    if (hasUpper && hasLeft && hasLength3 && hasAddressRef)
                    {
                        logicRows++;
                    }

                    var addressText = (ws.Cells[row, addressColIndex].Text ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(addressText))
                    {
                        continue;
                    }

                    var expected = (addressText.Length >= 3 ? addressText[..3] : addressText).ToUpperInvariant();
                    var actual = (postalCell.Text ?? string.Empty).Trim();
                    if (string.Equals(actual, expected, StringComparison.Ordinal))
                    {
                        valueRows++;
                    }
                }

                decimal score = 0;
                score += Math.Round(1m * formulaRows / rowCount, 2);
                score += Math.Round(2m * logicRows / rowCount, 2);
                score += Math.Round(1m * valueRows / rowCount, 2);

                if (logicRows != rowCount)
                {
                    result.Errors.Add($"Số dòng công thức Postal Code đúng chuẩn chưa đạt ({logicRows}/{rowCount}).");
                }
                if (valueRows != rowCount)
                {
                    result.Errors.Add($"Giá trị Postal Code uppercase chưa đạt ({valueRows}/{rowCount}).");
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

// minor-sync: non-functional graders update
