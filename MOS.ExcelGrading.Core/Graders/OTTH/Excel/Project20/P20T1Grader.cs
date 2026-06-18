using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project20
{
    public class P20T1Grader : ITaskGrader
    {
        public string TaskId => "P20-T1";
        public string TaskName => "Trong trang tính “London”, mở rộng công thức trong ô E5 xuống cuối cột của bảng.";
        public decimal MaxScore => 18m;

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
                var worksheet = P20GraderHelpers.GetSheet(studentSheet.Workbook, "London");
                if (worksheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'London'.");
                    return result;
                }

                var table = P20GraderHelpers.FindTable(
                    worksheet,
                    "Table1",
                    "Country or region",
                    "City",
                    "Air Miles",
                    "Bonus Air Miles");
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy bảng dữ liệu London có cột 'Air Miles' và 'Bonus Air Miles'.");
                    return result;
                }

                if (!P20GraderHelpers.TryGetColumnIndex(table, "Air Miles", out var airMilesCol)
                    || !P20GraderHelpers.TryGetColumnIndex(table, "Bonus Air Miles", out var bonusCol))
                {
                    result.Errors.Add("Không xác định được vị trí cột 'Air Miles' hoặc 'Bonus Air Miles'.");
                    return result;
                }

                decimal score = 0m;
                var startRow = table.Address.Start.Row + 1;
                var endRow = table.Address.End.Row;
                var totalRows = Math.Max(0, endRow - startRow + 1);
                if (totalRows == 0)
                {
                    result.Errors.Add("Bảng London không có dòng dữ liệu để chấm.");
                    return result;
                }

                var firstFormula = worksheet.Cells[startRow, bonusCol].Formula ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(firstFormula))
                {
                    score += 3m;
                    result.Details.Add("Ô E5 (dòng đầu cột Bonus Air Miles) đã có công thức.");
                }
                else
                {
                    result.Errors.Add("Ô E5 chưa có công thức gốc để mở rộng.");
                }

                var formulaRows = 0;
                var validFormulaRows = 0;
                var validValueRows = 0;

                for (var row = startRow; row <= endRow; row++)
                {
                    var formulaCell = worksheet.Cells[row, bonusCol];
                    var formula = formulaCell.Formula ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(formula))
                    {
                        continue;
                    }

                    formulaRows++;

                    var normalizedFormula = P20GraderHelpers.NormalizeFormula(formula);
                    var airMilesAddress = ExcelCellBase.GetAddress(row, airMilesCol).ToUpperInvariant();
                    var referencesAirMiles = normalizedFormula.Contains("AIRMILES", StringComparison.Ordinal)
                                             || normalizedFormula.Contains(airMilesAddress, StringComparison.Ordinal);
                    var hasExpectedRate = normalizedFormula.Contains("*0.08", StringComparison.Ordinal)
                                          || normalizedFormula.Contains("*8%", StringComparison.Ordinal);

                    if (referencesAirMiles && hasExpectedRate)
                    {
                        validFormulaRows++;
                    }

                    if (P20GraderHelpers.TryGetNumericValue(worksheet.Cells[row, airMilesCol], out var airMiles)
                        && P20GraderHelpers.TryGetNumericValue(formulaCell, out var bonusValue))
                    {
                        var expectedBonus = Math.Round(airMiles * 0.08m, 4, MidpointRounding.AwayFromZero);
                        if (Math.Abs(bonusValue - expectedBonus) <= 0.01m)
                        {
                            validValueRows++;
                        }
                    }
                }

                score += Math.Round(7m * formulaRows / totalRows, 2, MidpointRounding.AwayFromZero);
                if (formulaRows == totalRows)
                {
                    result.Details.Add("Công thức đã được mở rộng đủ toàn bộ cột Bonus Air Miles trong bảng.");
                }
                else
                {
                    result.Errors.Add($"Công thức chưa được mở rộng đầy đủ trong bảng ({formulaRows}/{totalRows} dòng).");
                }

                score += Math.Round(4m * validFormulaRows / totalRows, 2, MidpointRounding.AwayFromZero);
                if (validFormulaRows == totalRows)
                {
                    result.Details.Add("Công thức ở cột Bonus Air Miles tham chiếu đúng cột Air Miles và đúng tỷ lệ 8%.");
                }
                else
                {
                    result.Errors.Add($"Có dòng dùng công thức sai logic ở cột Bonus Air Miles ({validFormulaRows}/{totalRows} dòng đúng).");
                }

                score += Math.Round(2m * validValueRows / totalRows, 2, MidpointRounding.AwayFromZero);
                if (validValueRows == totalRows)
                {
                    result.Details.Add("Giá trị tính ra ở cột Bonus Air Miles khớp với phép nhân Air Miles × 8%.");
                }
                else
                {
                    result.Errors.Add($"Kết quả số trong cột Bonus Air Miles chưa chính xác ở một số dòng ({validValueRows}/{totalRows} dòng đúng).");
                }

                var extraFormulaRows = new List<int>();
                for (var row = endRow + 1; row <= Math.Min(worksheet.Dimension.End.Row, endRow + 10); row++)
                {
                    if (!string.IsNullOrWhiteSpace(worksheet.Cells[row, bonusCol].Formula))
                    {
                        extraFormulaRows.Add(row);
                    }
                }

                if (extraFormulaRows.Count == 0)
                {
                    score += 2m;
                    result.Details.Add("Không có công thức bị kéo thừa ra ngoài vùng dữ liệu của bảng.");
                }
                else
                {
                    result.Errors.Add($"Có công thức kéo thừa ngoài bảng tại các dòng: {string.Join(", ", extraFormulaRows)}.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 1: {ex.Message}.");
            }

            return result;
        }
    }
}


