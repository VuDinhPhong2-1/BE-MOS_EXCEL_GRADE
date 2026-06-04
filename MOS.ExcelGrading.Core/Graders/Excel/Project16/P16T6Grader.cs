using System.Xml;
using System.Text.RegularExpressions;
using System.Globalization;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.Project16
{
    public class P16T6Grader : ITaskGrader
    {
        public string TaskId => "P16-T6";
        public string TaskName => "Summary: bieu do Colorful Palette 2";
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
                var ws = P16GraderHelpers.GetSheet(studentSheet.Workbook, "Summary");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Summary'.");
                    return result;
                }

                var drawing = ws.Drawings.FirstOrDefault();
                if (drawing == null)
                {
                    result.Errors.Add("Không tìm thấy biểu đồ tren sheet 'Summary'.");
                    return result;
                }

                decimal score = 0m;
                score += 1m;
                result.Details.Add("Tìm thấy chart tren sheet 'Summary'.");

                var (colorStyleId, chartStyleId) = P16GraderHelpers.GetDrawingStyleIds(drawing);
                if (string.Equals(colorStyleId, "11", StringComparison.Ordinal))
                {
                    score += 2m;
                    result.Details.Add("Color style ID = 11 (Colorful Palette 2).");
                }
                else
                {
                    result.Errors.Add($"Color style ID chưa đúng. Hiện tại: '{colorStyleId}' (mong đợi 11).");
                }

                if (string.Equals(chartStyleId, "410", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Chart style ID dung: 410.");
                }
                else
                {
                    result.Errors.Add($"Chart style ID chưa đúng. Hiện tại: '{chartStyleId}' (mong đợi 410).");
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



