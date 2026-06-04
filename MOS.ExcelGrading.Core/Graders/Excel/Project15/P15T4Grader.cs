using System.Xml;
using System.Text.RegularExpressions;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project15
{
    public class P15T4Grader : ITaskGrader
    {
        public string TaskId => "P15-T4";
        public string TaskName => "Customers: cong thuc N5 COUNTIF United States";
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
                var ws = P15GraderHelpers.GetSheet(studentSheet.Workbook, "Customers");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Customers'.");
                    return result;
                }

                var actual = P15GraderHelpers.NormalizeFormula(ws.Cells["N5"].Formula);
                var expected = P15GraderHelpers.NormalizeFormula("COUNTIF(I4:I32,\"United States\")");
                if (string.Equals(actual, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Công thức N5 dung chinh xac.");
                }
                else
                {
                    result.Errors.Add($"Công thức N5 chưa đúng. Hiện tại: '{ws.Cells["N5"].Formula}'.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi: {ex.Message}");
            }

            return result;
        }
    }
}



