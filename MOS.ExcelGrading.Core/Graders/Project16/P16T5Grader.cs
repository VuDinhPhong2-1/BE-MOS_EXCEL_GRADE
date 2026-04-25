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
    public class P16T5Grader : ITaskGrader
    {
        public string TaskId => "P16-T5";
        public string TaskName => "Products: cong thuc Estimated Value tai F3 va Fill Down";
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
                var ws = P16GraderHelpers.GetSheet(studentSheet.Workbook, "Products");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Products'.");
                    return result;
                }

                decimal score = 0m;
                var formulaF3 = P16GraderHelpers.NormalizeFormula(ws.Cells["F3"].Formula);
                var expectedF3 = P16GraderHelpers.NormalizeFormula("D3*E3");
                if (string.Equals(formulaF3, expectedF3, StringComparison.Ordinal))
                {
                    score += 2m;
                    result.Details.Add("Công thức gốc F3 dung: D3*E3.");
                }
                else
                {
                    result.Errors.Add($"Công thức F3 chưa đúng. Hiện tại: '{ws.Cells["F3"].Formula}'.");
                }

                var filledRows = 0;
                for (var row = 3; row <= 34; row++)
                {
                    if (!string.IsNullOrWhiteSpace(ws.Cells[row, 6].Formula))
                    {
                        filledRows++;
                    }
                }

                if (filledRows == 32)
                {
                    score += 2m;
                    result.Details.Add("Công thức da duoc fill down day du F3:F34.");
                }
                else
                {
                    result.Errors.Add($"Công thức chua fill day du F3:F34 (hien tai {filledRows}/32).");
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



