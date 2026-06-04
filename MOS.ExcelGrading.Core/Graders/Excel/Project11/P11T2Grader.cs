using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project11
{
    public class P11T2Grader : ITaskGrader
    {
        public string TaskId => "P11-T2";
        public string TaskName => "Shareholders Info: chieu cao dong Annual Report = 30";
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
                var ws = P11GraderHelpers.GetSheet(studentSheet.Workbook, "Shareholders Info");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Shareholders Info'.");
                    return result;
                }

                decimal score = 0m;
                var annualRowIndex = -1;
                for (var row = 1; row <= ws.Dimension.End.Row; row++)
                {
                    var rowText = ws.Cells[row, 1, row, ws.Dimension.End.Column].Text;
                    if (!string.IsNullOrWhiteSpace(rowText)
                        && rowText.Contains("Annual Report", StringComparison.OrdinalIgnoreCase))
                    {
                        annualRowIndex = row;
                        break;
                    }
                }

                if (annualRowIndex < 0)
                {
                    result.Errors.Add("Không tìm thấy dòng chua giá trị 'Annual Report'.");
                    return result;
                }

                score += 1m;
                result.Details.Add($"Tìm thấy dòng 'Annual Report' tai hàng {annualRowIndex}.");

                var targetRow = ws.Row(annualRowIndex);
                if (Math.Abs(targetRow.Height - 30d) <= 0.1d)
                {
                    score += 3m;
                    result.Details.Add("Do cao dòng dung 30.");
                }
                else
                {
                    result.Errors.Add($"Do cao dòng chưa đúng. Hiện tại: {targetRow.Height:0.##}, mong đợi: 30.");
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



