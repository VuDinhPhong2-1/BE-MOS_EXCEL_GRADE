using System.Xml;
using System.Text.RegularExpressions;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project15
{
    public class P15T2Grader : ITaskGrader
    {
        public string TaskId => "P15-T2";
        public string TaskName => "Products: cong thuc G3 SUMIF cho Magic Supplies";
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
                var ws = P15GraderHelpers.GetSheet(studentSheet.Workbook, "Products");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Products'.");
                    return result;
                }

                var actual = P15GraderHelpers.NormalizeFormula(ws.Cells["G3"].Formula);
                var expected = P15GraderHelpers.NormalizeFormula("SUMIF(Table2[Catergory],\"Magic Supplies\",Table2[Weight])");
                if (string.Equals(actual, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Công thức G3 dung chinh xac.");
                }
                else
                {
                    result.Errors.Add($"Công thức G3 chưa đúng. Hiện tại: '{ws.Cells["G3"].Formula}'.");
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



