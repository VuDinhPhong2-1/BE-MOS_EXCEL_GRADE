using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project18
{
    public class P18T2Grader : ITaskGrader
    {
        public string TaskId => "P18-T2";
        public string TaskName => "Trong trang tính \"Exchange Rates\", định dạng các ô từ B4 đến D8 để hiển thị số với tối đa hai chữ số thập phân.";
        public decimal MaxScore => 20m;

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
                var worksheet = P18GraderHelpers.GetSheet(studentSheet.Workbook, "Exchange Rates");
                if (worksheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Exchange Rates'.");
                    return result;
                }

                decimal score = 0m;

                var targetRange = worksheet.Cells["B4:D8"];
                var targetTotal = 0;
                var targetPass = 0;
                var wrongTargetCells = new List<string>();

                foreach (var cell in targetRange)
                {
                    targetTotal++;
                    if (P18GraderHelpers.IsTwoDecimalNumberFormat(cell))
                    {
                        targetPass++;
                    }
                    else
                    {
                        wrongTargetCells.Add(cell.Address);
                    }
                }

                if (targetTotal > 0)
                {
                    var coverageScore = Math.Round(12m * targetPass / targetTotal, 2, MidpointRounding.AwayFromZero);
                    score += coverageScore;

                    if (targetPass == targetTotal)
                    {
                        result.Details.Add("Phạm vi B4:D8 đã được định dạng đủ 2 chữ số thập phân.");
                    }
                    else
                    {
                        result.Errors.Add(
                            $"Một số ô trong B4:D8 chưa đúng định dạng 2 chữ số thập phân ({targetPass}/{targetTotal}). Ô lỗi: {string.Join(", ", wrongTargetCells)}.");
                    }
                }
                else
                {
                    result.Errors.Add("Không đọc được phạm vi mục tiêu B4:D8.");
                }

                var guardRanges = new[] { "B3:D3", "A4:A8", "B9:D9" };
                var guardTotal = 0;
                var guardPass = 0;
                var overAppliedCells = new List<string>();

                foreach (var guardRange in guardRanges)
                {
                    foreach (var cell in worksheet.Cells[guardRange])
                    {
                        guardTotal++;
                        if (!P18GraderHelpers.IsTwoDecimalNumberFormat(cell))
                        {
                            guardPass++;
                        }
                        else
                        {
                            overAppliedCells.Add(cell.Address);
                        }
                    }
                }

                if (guardTotal > 0)
                {
                    var scopeScore = Math.Round(8m * guardPass / guardTotal, 2, MidpointRounding.AwayFromZero);
                    score += scopeScore;

                    if (guardPass == guardTotal)
                    {
                        result.Details.Add("Phạm vi áp dụng định dạng là đúng, không bị thừa sang vùng lân cận.");
                    }
                    else
                    {
                        result.Errors.Add(
                            $"Định dạng 2 chữ số thập phân bị áp dư ra ngoài vùng yêu cầu ({guardPass}/{guardTotal} ô vùng kiểm tra đạt). Ô bị áp dư: {string.Join(", ", overAppliedCells)}.");
                    }
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 2: {ex.Message}.");
            }

            return result;
        }
    }
}


