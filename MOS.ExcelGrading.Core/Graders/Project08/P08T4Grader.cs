using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project08
{
    public class P08T4Grader : ITaskGrader
    {
        public string TaskId => "P08-T4";
        public string TaskName => "Authors Premium: IF(Books sold > 10000, 500, 100)";
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
                var ws = P08GraderHelpers.GetSheet(studentSheet, "Authors");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Authors'.");
                    return result;
                }

                var table = ws.Tables.FirstOrDefault(t =>
                    t.Columns.Any(c => string.Equals((c.Name ?? string.Empty).Trim(), "Books sold", StringComparison.OrdinalIgnoreCase)) &&
                    t.Columns.Any(c => string.Equals((c.Name ?? string.Empty).Trim(), "Premium", StringComparison.OrdinalIgnoreCase)));
                if (table == null)
                {
                    result.Errors.Add("Khong tim thay table co cot 'Books sold' va 'Premium'.");
                    return result;
                }

                var booksSoldCol = table.Columns.First(c =>
                    string.Equals((c.Name ?? string.Empty).Trim(), "Books sold", StringComparison.OrdinalIgnoreCase));
                var premiumCol = table.Columns.First(c =>
                    string.Equals((c.Name ?? string.Empty).Trim(), "Premium", StringComparison.OrdinalIgnoreCase));

                var startRow = table.Address.Start.Row + 1;
                var endRow = table.Address.End.Row - (table.ShowTotal ? 1 : 0);
                var rowCount = Math.Max(0, endRow - startRow + 1);
                if (rowCount == 0)
                {
                    result.Errors.Add("Khong co du lieu de cham.");
                    return result;
                }

                var booksColIndex = table.Address.Start.Column + booksSoldCol.Position;
                var premiumColIndex = table.Address.Start.Column + premiumCol.Position;

                var formulaRows = 0;
                var logicRows = 0;
                var valueRows = 0;

                for (var row = startRow; row <= endRow; row++)
                {
                    var premiumCell = ws.Cells[row, premiumColIndex];
                    var formulaRaw = premiumCell.Formula ?? string.Empty;
                    var formula = P08GraderHelpers.NormalizeFormula(formulaRaw);

                    if (!string.IsNullOrWhiteSpace(formulaRaw))
                    {
                        formulaRows++;
                    }
                    else
                    {
                        result.Errors.Add($"Hang {row}: Premium chua co cong thuc.");
                        continue;
                    }

                    var booksAddress = $"{OfficeOpenXml.ExcelCellBase.GetAddress(row, booksColIndex)}";
                    var hasBooksRef =
                        formula.Contains("BOOKSSOLD", StringComparison.Ordinal) ||
                        formula.Contains(booksAddress, StringComparison.OrdinalIgnoreCase);
                    var hasIf = formula.Contains("IF(", StringComparison.Ordinal);
                    var hasThreshold = formula.Contains(">10000", StringComparison.Ordinal);
                    var hasTrueFalse = formula.Contains(",500,100", StringComparison.Ordinal) ||
                                       formula.Contains(",500;100", StringComparison.Ordinal);

                    if (hasIf && hasThreshold && hasTrueFalse && hasBooksRef)
                    {
                        logicRows++;
                    }
                    else
                    {
                        result.Errors.Add($"Hang {row}: Cong thuc Premium chua dung logic IF >10000,500,100.");
                    }

                    var sold = P08GraderHelpers.ToDecimal(ws.Cells[row, booksColIndex].Value, ws.Cells[row, booksColIndex].Text);
                    var expected = sold > 10000m ? 500m : 100m;
                    var actual = P08GraderHelpers.ToDecimal(premiumCell.Value, premiumCell.Text);
                    if (Math.Abs(expected - actual) <= 0.01m)
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
                    result.Errors.Add($"So dong cong thuc Premium dung chuan chua dat ({logicRows}/{rowCount}).");
                }
                if (valueRows != rowCount)
                {
                    result.Errors.Add($"Gia tri Premium tinh dung chua dat ({valueRows}/{rowCount}).");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }
}
