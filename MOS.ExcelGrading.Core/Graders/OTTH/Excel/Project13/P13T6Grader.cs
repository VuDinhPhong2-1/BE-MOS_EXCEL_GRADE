using System.Globalization;
using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project13
{
    public class P13T6Grader : ITaskGrader
    {
        public string TaskId => "P13-T6";
        public string TaskName => "Price List: cong thuc H5 dung Structured Reference";
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
                var ws = P13GraderHelpers.GetSheet(studentSheet.Workbook, "Price List");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'Price List'.");
                    return result;
                }

                var actual = P13GraderHelpers.NormalizeFormula(ws.Cells["H5"].Formula);
                var expected = P13GraderHelpers.NormalizeFormula("ROWS(Phones[])");
                if (string.Equals(actual, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Công th?c H5 dung: ROWS(Phones[]).");
                }
                else
                {
                    result.Errors.Add($"Công th?c H5 chua dúng. Hi?n t?i: '{ws.Cells["H5"].Formula}'.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"L?i: {ex.Message}");
            }

            return result;
        }
    }
}




