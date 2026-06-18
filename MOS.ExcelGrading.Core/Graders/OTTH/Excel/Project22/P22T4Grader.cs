using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project22
{
    public class P22T4Grader : ITaskGrader
    {
        public string TaskId => "P22-T4";
        public string TaskName => "Trên trang tính \"Scoring Criteria\", tại ô B28, nhập công thức để cộng các giá trị trong các phạm vi đã đặt tên là \"Total 1\", \"Total 2\" và \"Total 3\". Sử dụng tên phạm vi trong công thức thay vì tham chiếu ô hoặc giá trị cụ thể.";
        public decimal MaxScore => 24m;

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
                var worksheet = P22GraderHelpers.GetSheet(studentSheet.Workbook, "Scoring Criteria");
                if (worksheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Scoring Criteria'.");
                    return result;
                }

                decimal score = 0m;

                var targetCell = worksheet.Cells["B28"];
                var rawFormula = targetCell.Formula ?? string.Empty;
                var normalizedFormula = P22GraderHelpers.NormalizeFormula(rawFormula);

                if (!string.IsNullOrWhiteSpace(rawFormula))
                {
                    score += 5m;
                    result.Details.Add("Ô B28 đã có công thức.");
                }
                else
                {
                    result.Errors.Add("Ô B28 chưa có công thức.");
                    result.Score = score;
                    return result;
                }

                if (normalizedFormula.StartsWith("SUM(", StringComparison.Ordinal)
                    && normalizedFormula.EndsWith(")", StringComparison.Ordinal))
                {
                    score += 5m;
                    result.Details.Add("Công thức tại B28 đã dùng hàm SUM.");
                }
                else
                {
                    result.Errors.Add($"Công thức tại B28 chưa dùng hàm SUM. Hiện tại: '{rawFormula}'.");
                }

                var hasTotal1 = normalizedFormula.Contains("TOTAL1", StringComparison.Ordinal);
                var hasTotal2 = normalizedFormula.Contains("TOTAL2", StringComparison.Ordinal);
                var hasTotal3 = normalizedFormula.Contains("TOTAL3", StringComparison.Ordinal);
                if (hasTotal1 && hasTotal2 && hasTotal3)
                {
                    score += 6m;
                    result.Details.Add("Công thức đã tham chiếu đầy đủ các tên phạm vi Total1, Total2 và Total3.");
                }
                else
                {
                    result.Errors.Add(
                        $"Công thức chưa tham chiếu đầy đủ tên phạm vi. Total1={hasTotal1}, Total2={hasTotal2}, Total3={hasTotal3}.");
                }

                var referencesCellAddress = normalizedFormula.Contains("B16", StringComparison.Ordinal)
                                            || normalizedFormula.Contains("B22", StringComparison.Ordinal)
                                            || normalizedFormula.Contains("B27", StringComparison.Ordinal);
                if (!referencesCellAddress)
                {
                    score += 4m;
                    result.Details.Add("Công thức không dùng tham chiếu ô trực tiếp (B16, B22, B27).");
                }
                else
                {
                    result.Errors.Add("Công thức đang dùng tham chiếu ô trực tiếp thay vì dùng tên phạm vi.");
                }

                var argumentText = normalizedFormula;
                if (argumentText.StartsWith("SUM(", StringComparison.Ordinal)
                    && argumentText.EndsWith(")", StringComparison.Ordinal))
                {
                    argumentText = argumentText[4..^1];
                }

                var arguments = argumentText
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(part => part.Trim())
                    .ToList();

                var hasExactThreeNamedArguments = arguments.Count == 3
                                                  && arguments.All(part =>
                                                      string.Equals(part, "TOTAL1", StringComparison.OrdinalIgnoreCase)
                                                      || string.Equals(part, "TOTAL2", StringComparison.OrdinalIgnoreCase)
                                                      || string.Equals(part, "TOTAL3", StringComparison.OrdinalIgnoreCase))
                                                  && arguments.Distinct(StringComparer.OrdinalIgnoreCase).Count() == 3;

                if (hasExactThreeNamedArguments)
                {
                    score += 2m;
                    result.Details.Add("Cấu trúc tham số của hàm SUM đúng theo dạng SUM(Total1,Total2,Total3).");
                }
                else
                {
                    result.Errors.Add($"Cấu trúc tham số của hàm SUM chưa chính xác. Hiện tại: '{rawFormula}'.");
                }

                var n1 = worksheet.Workbook.Names.FirstOrDefault(name =>
                    string.Equals(name.Name, "Total1", StringComparison.OrdinalIgnoreCase));
                var n2 = worksheet.Workbook.Names.FirstOrDefault(name =>
                    string.Equals(name.Name, "Total2", StringComparison.OrdinalIgnoreCase));
                var n3 = worksheet.Workbook.Names.FirstOrDefault(name =>
                    string.Equals(name.Name, "Total3", StringComparison.OrdinalIgnoreCase));

                decimal v1 = 0m;
                decimal v2 = 0m;
                decimal v3 = 0m;
                decimal actual = 0m;
                var canCheckValue = n1 != null
                                    && n2 != null
                                    && n3 != null
                                    && P22GraderHelpers.TryGetNumericValue(n1, out v1)
                                    && P22GraderHelpers.TryGetNumericValue(n2, out v2)
                                    && P22GraderHelpers.TryGetNumericValue(n3, out v3)
                                    && P22GraderHelpers.TryGetNumericValue(targetCell, out actual);

                if (canCheckValue)
                {
                    var expected = v1 + v2 + v3;
                    if (Math.Abs(expected - actual) <= 0.01m)
                    {
                        score += 2m;
                        result.Details.Add("Giá trị kết quả tại B28 khớp với tổng của Total1, Total2 và Total3.");
                    }
                    else
                    {
                        result.Errors.Add(
                            $"Giá trị kết quả tại B28 chưa đúng. Hiện tại: {actual}, mong đợi: {expected}.");
                    }
                }
                else
                {
                    result.Errors.Add("Không đủ điều kiện để đối chiếu giá trị kết quả với các tên phạm vi.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 4: {ex.Message}.");
            }

            return result;
        }
    }
}

