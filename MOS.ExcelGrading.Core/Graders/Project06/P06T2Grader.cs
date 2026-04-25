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
                    result.Errors.Add("Không tìm thấy sheet 'Region 1'.");
                    return result;
                }

                var table = ws.Tables.FirstOrDefault();
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy bảng dữ liệu trên sheet 'Region 1'.");
                    return result;
                }

                var productCol = table.Columns.FirstOrDefault(c =>
                    string.Equals((c.Name ?? string.Empty).Trim(), "Product", StringComparison.OrdinalIgnoreCase));
                var totalSalesCol = table.Columns.FirstOrDefault(c =>
                    string.Equals((c.Name ?? string.Empty).Trim(), "Total sales", StringComparison.OrdinalIgnoreCase));
                if (productCol == null || totalSalesCol == null)
                {
                    result.Errors.Add("Thiếu cột Product hoặc Total sales trong bảng 'Region 1'.");
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
                    result.Errors.Add("Không có dòng dữ liệu để kiểm tra sort.");
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
                    result.Details.Add("Dữ liệu đã được sort đúng Product A-Z và Total sales giảm dần.");
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
                            $"Thứ tự sort sai tại vị trí {mismatchIndex + 1}: hiện tại '{actual.Product}' ({actual.Sales}), mong đợi '{exp.Product}' ({exp.Sales}).");
                    }
                    else
                    {
                        result.Errors.Add("Thứ tự sort chưa đúng theo yêu cầu multi-level.");
                    }
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
