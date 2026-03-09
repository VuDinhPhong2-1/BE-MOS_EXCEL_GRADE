using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project06
{
    public class P06T2Grader : ITaskGrader
    {
        public string TaskId => "P06-T2";
        public string TaskName => "Region 1: Multi-level sort Product A-Z, then Total sales desc";
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
                var ws = P06GraderHelpers.GetSheet(studentSheet, "Region 1");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Region 1'.");
                    return result;
                }

                var table = ws.Tables.FirstOrDefault();
                if (table == null)
                {
                    result.Errors.Add("Khong tim thay bang du lieu tren Region 1.");
                    return result;
                }

                var productCol = table.Columns.FirstOrDefault(c =>
                    string.Equals((c.Name ?? string.Empty).Trim(), "Product", StringComparison.OrdinalIgnoreCase));
                var totalSalesCol = table.Columns.FirstOrDefault(c =>
                    string.Equals((c.Name ?? string.Empty).Trim(), "Total sales", StringComparison.OrdinalIgnoreCase));
                if (productCol == null || totalSalesCol == null)
                {
                    result.Errors.Add("Thieu cot Product hoac Total sales trong table Region 1.");
                    return result;
                }

                decimal score = 1m; // Tim thay table + 2 cot can cham.
                var startRow = table.Address.Start.Row + 1;
                var endRow = table.Address.End.Row - (table.ShowTotal ? 1 : 0);
                var productColIndex = table.Address.Start.Column + productCol.Position;
                var totalSalesColIndex = table.Address.Start.Column + totalSalesCol.Position;

                var rows = new List<(string Product, decimal Sales)>();
                for (var row = startRow; row <= endRow; row++)
                {
                    var product = (ws.Cells[row, productColIndex].Text ?? string.Empty).Trim();
                    var sales = P06GraderHelpers.ToDecimal(
                        ws.Cells[row, totalSalesColIndex].Value,
                        ws.Cells[row, totalSalesColIndex].Text);
                    if (string.IsNullOrWhiteSpace(product))
                    {
                        continue;
                    }
                    rows.Add((product, sales));
                }

                if (rows.Count == 0)
                {
                    result.Errors.Add("Khong co dong du lieu de cham sort.");
                    result.Score = score;
                    return result;
                }

                var expected = rows
                    .OrderBy(r => r.Product, StringComparer.OrdinalIgnoreCase)
                    .ThenByDescending(r => r.Sales)
                    .ToList();

                var exactOrderMatch = rows.SequenceEqual(expected);
                if (exactOrderMatch)
                {
                    score += 3m;
                    result.Details.Add("Du lieu da duoc sort dung Product A-Z va Total sales giam dan.");
                }
                else
                {
                    var mismatchIndex = -1;
                    for (var i = 0; i < rows.Count; i++)
                    {
                        var row = rows[i];
                        var expRow = expected[i];
                        if (!string.Equals(row.Product, expRow.Product, StringComparison.OrdinalIgnoreCase) ||
                            row.Sales != expRow.Sales)
                        {
                            mismatchIndex = i;
                            break;
                        }
                    }
                    if (mismatchIndex >= 0)
                    {
                        var actual = rows[mismatchIndex];
                        var exp = expected[mismatchIndex];
                        result.Errors.Add(
                            $"Thu tu sort sai tai vi tri {mismatchIndex + 1}: hien tai '{actual.Product}' ({actual.Sales}), mong doi '{exp.Product}' ({exp.Sales}).");
                    }
                    else
                    {
                        result.Errors.Add("Thu tu sort chua dung theo yeu cau multi-level.");
                    }
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
