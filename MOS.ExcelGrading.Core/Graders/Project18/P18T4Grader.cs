using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project18
{
    public class P18T4Grader : ITaskGrader
    {
        public string TaskId => "P18-T4";
        public string TaskName => "Trong trang tính \"Key Accounts\", tại cột Monthly Average, sử dụng hàm để tính số dư trung bình hàng tháng cho mỗi tài khoản, dựa trên dữ liệu từ tháng 1 đến tháng 4.";
        public decimal MaxScore => 20m;

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
                var worksheet = P18GraderHelpers.GetSheet(studentSheet.Workbook, "Key Accounts");
                if (worksheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Key Accounts'.");
                    return result;
                }

                decimal score = 0m;

                var table = P18GraderHelpers.FindTableByHeaders(
                    worksheet,
                    "Account Name",
                    "Monthly Average",
                    "January",
                    "February",
                    "March",
                    "April");
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy bảng dữ liệu có cột Monthly Average và các tháng 1 đến 4.");
                    return result;
                }

                score += 4m;
                result.Details.Add("Đã tìm thấy bảng dữ liệu Key Accounts với đầy đủ cột yêu cầu.");

                var columnsByNormalizedName = table.Columns
                    .ToDictionary(
                        column => P18GraderHelpers.NormalizeIdentifier(column.Name),
                        column => table.Address.Start.Column + column.Position,
                        StringComparer.OrdinalIgnoreCase);

                if (!columnsByNormalizedName.TryGetValue("MONTHLYAVERAGE", out var monthlyAverageColumn)
                    || !columnsByNormalizedName.TryGetValue("JANUARY", out var januaryColumn)
                    || !columnsByNormalizedName.TryGetValue("FEBRUARY", out var februaryColumn)
                    || !columnsByNormalizedName.TryGetValue("MARCH", out var marchColumn)
                    || !columnsByNormalizedName.TryGetValue("APRIL", out var aprilColumn))
                {
                    result.Errors.Add("Không xác định được đúng cột Monthly Average hoặc các cột tháng 1 đến 4.");
                    result.Score = score;
                    return result;
                }

                var monthColumns = new[] { januaryColumn, februaryColumn, marchColumn, aprilColumn };
                var firstDataRow = table.Address.Start.Row + 1;
                var lastDataRow = table.Address.End.Row;
                var totalRows = Math.Max(0, lastDataRow - firstDataRow + 1);
                if (totalRows == 0)
                {
                    result.Errors.Add("Bảng Key Accounts không có dòng dữ liệu để chấm.");
                    result.Score = score;
                    return result;
                }

                var formulaPresentCount = 0;
                var formulaLogicCount = 0;
                var numericResultCount = 0;

                for (var row = firstDataRow; row <= lastDataRow; row++)
                {
                    var cell = worksheet.Cells[row, monthlyAverageColumn];
                    var formula = cell.Formula ?? string.Empty;

                    if (!string.IsNullOrWhiteSpace(formula))
                    {
                        formulaPresentCount++;

                        if (P18GraderHelpers.IsAverageFormulaForMonthlyRange(formula, row, januaryColumn, aprilColumn))
                        {
                            formulaLogicCount++;
                        }
                    }

                    var allMonthValuesAreNumeric = true;
                    var monthValues = new List<decimal>();
                    foreach (var monthColumn in monthColumns)
                    {
                        if (P18GraderHelpers.TryGetNumericValue(worksheet.Cells[row, monthColumn], out var monthValue))
                        {
                            monthValues.Add(monthValue);
                        }
                        else
                        {
                            allMonthValuesAreNumeric = false;
                            break;
                        }
                    }

                    if (!allMonthValuesAreNumeric || monthValues.Count != 4)
                    {
                        continue;
                    }

                    var expectedAverage = monthValues.Average();
                    if (P18GraderHelpers.TryGetNumericValue(cell, out var actualAverage))
                    {
                        if (Math.Abs(actualAverage - expectedAverage) <= 0.01m)
                        {
                            numericResultCount++;
                        }
                    }
                }

                score += Math.Round(4m * formulaPresentCount / totalRows, 2, MidpointRounding.AwayFromZero);
                if (formulaPresentCount == totalRows)
                {
                    result.Details.Add("Cột Monthly Average đã có công thức cho toàn bộ dòng dữ liệu.");
                }
                else
                {
                    result.Errors.Add(
                        $"Cột Monthly Average chưa có công thức đầy đủ ({formulaPresentCount}/{totalRows} dòng).");
                }

                score += Math.Round(8m * formulaLogicCount / totalRows, 2, MidpointRounding.AwayFromZero);
                if (formulaLogicCount == totalRows)
                {
                    result.Details.Add("Công thức Monthly Average dùng hàm AVERAGE và tham chiếu đúng dữ liệu tháng 1 đến tháng 4.");
                }
                else
                {
                    result.Errors.Add(
                        $"Công thức Monthly Average chưa đúng phạm vi tháng 1 đến tháng 4 ({formulaLogicCount}/{totalRows} dòng đúng).");
                }

                score += Math.Round(4m * numericResultCount / totalRows, 2, MidpointRounding.AwayFromZero);
                if (numericResultCount == totalRows)
                {
                    result.Details.Add("Kết quả số tại cột Monthly Average khớp với trung bình của 4 tháng.");
                }
                else
                {
                    result.Errors.Add(
                        $"Giá trị tính trung bình tại cột Monthly Average chưa chính xác ở một số dòng ({numericResultCount}/{totalRows} dòng đúng).");
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

