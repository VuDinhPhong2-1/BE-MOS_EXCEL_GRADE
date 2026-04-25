using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.Project12
{
    public class P12T2Grader : ITaskGrader
    {
        public string TaskId => "P12-T2";
        public string TaskName => "Prices: ap dung Cell Style Title cho A1";
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
                var ws = P12GraderHelpers.GetSheet(studentSheet.Workbook, "Prices");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Prices'.");
                    return result;
                }

                var cell = ws.Cells["A1"];
                decimal score = 0m;

                if (string.Equals(cell.Style.Font.Name ?? string.Empty, "Century Gothic", StringComparison.OrdinalIgnoreCase))
                {
                    score += 1m;
                    result.Details.Add("Font A1 dung: Century Gothic.");
                }
                else
                {
                    result.Errors.Add($"Font A1 chưa đúng. Hiện tại: '{cell.Style.Font.Name}'.");
                }

                if (Math.Abs(cell.Style.Font.Size - 18f) <= 0.1f)
                {
                    score += 1m;
                    result.Details.Add("Co chu A1 dung: 18.");
                }
                else
                {
                    result.Errors.Add($"Co chu A1 chưa đúng. Hiện tại: {cell.Style.Font.Size:0.##}.");
                }

                if (cell.Style.HorizontalAlignment == ExcelHorizontalAlignment.Left)
                {
                    score += 1m;
                    result.Details.Add("Can le ngang A1 dung: Left.");
                }
                else
                {
                    result.Errors.Add($"Can le ngang A1 chưa đúng. Hiện tại: {cell.Style.HorizontalAlignment}.");
                }

                if (cell.Style.VerticalAlignment == ExcelVerticalAlignment.Center)
                {
                    score += 1m;
                    result.Details.Add("Can le doc A1 dung: Center.");
                }
                else
                {
                    result.Errors.Add($"Can le doc A1 chưa đúng. Hiện tại: {cell.Style.VerticalAlignment}.");
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



