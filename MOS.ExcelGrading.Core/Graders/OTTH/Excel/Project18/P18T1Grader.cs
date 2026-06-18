using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project18
{
    public class P18T1Grader : ITaskGrader
    {
        public string TaskId => "P18-T1";
        public string TaskName => "Đi tới phạm vi đã đặt tên Rate và xóa nội dung của các ô trong phạm vi đó.";
        public decimal MaxScore => 15m;

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
                var workbook = studentSheet.Workbook;
                var exchangeRatesSheet = P18GraderHelpers.GetSheet(workbook, "Exchange Rates");
                if (exchangeRatesSheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Exchange Rates'.");
                    return result;
                }

                decimal score = 0m;
                var namedRange = P18GraderHelpers.FindNamedRange(workbook, "Rate");
                var targetAddress = "A11:B11";

                if (namedRange != null)
                {
                    var normalizedAddress = P18GraderHelpers.NormalizeRange(namedRange.FullAddressAbsolute);
                    var expectedAddress = P18GraderHelpers.NormalizeRange(targetAddress);
                    var rangeOnExpectedSheet = string.Equals(
                        (namedRange.Worksheet?.Name ?? exchangeRatesSheet.Name).Trim(),
                        exchangeRatesSheet.Name.Trim(),
                        StringComparison.OrdinalIgnoreCase);

                    if (rangeOnExpectedSheet && string.Equals(normalizedAddress, expectedAddress, StringComparison.OrdinalIgnoreCase))
                    {
                        score += 5m;
                        result.Details.Add("Named range 'Rate' tồn tại và vẫn trỏ đúng đến 'Exchange Rates'!A11:B11.");
                    }
                    else
                    {
                        result.Errors.Add(
                            $"Named range 'Rate' chưa đúng phạm vi. Hiện tại: '{namedRange.FullAddressAbsolute}'.");
                    }

                    targetAddress = normalizedAddress;
                }
                else
                {
                    result.Errors.Add("Không tìm thấy named range 'Rate'.");
                }

                var cells = exchangeRatesSheet.Cells[targetAddress];
                var totalCells = 0;
                var emptyCells = 0;
                var unclearedCells = new List<string>();

                foreach (var cell in cells)
                {
                    totalCells++;
                    if (P18GraderHelpers.CellIsEmpty(cell))
                    {
                        emptyCells++;
                    }
                    else
                    {
                        unclearedCells.Add(cell.Address);
                    }
                }

                if (totalCells == 0)
                {
                    result.Errors.Add("Không xác định được ô nào trong phạm vi cần kiểm tra.");
                    result.Score = score;
                    return result;
                }

                if (emptyCells == totalCells)
                {
                    score += 10m;
                    result.Details.Add("Toàn bộ ô trong phạm vi Rate đã được xóa nội dung hoàn toàn.");
                }
                else
                {
                    var partial = Math.Round(10m * emptyCells / totalCells, 2, MidpointRounding.AwayFromZero);
                    score += partial;
                    result.Errors.Add(
                        $"Phạm vi Rate chưa được xóa sạch nội dung ({emptyCells}/{totalCells}). Ô còn dữ liệu: {string.Join(", ", unclearedCells)}.");
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


