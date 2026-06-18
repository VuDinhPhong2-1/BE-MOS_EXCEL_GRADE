using System.Xml;
using System.Text.RegularExpressions;
using System.Globalization;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project16
{
    public class P16T2Grader : ITaskGrader
    {
        public string TaskId => "P16-T2";
        public string TaskName => "Products: can trai o A1";
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
                var ws = P16GraderHelpers.GetSheet(studentSheet.Workbook, "Products");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'Products'.");
                    return result;
                }

                if (ws.Cells["A1"].Style.HorizontalAlignment == ExcelHorizontalAlignment.Left)
                {
                    result.Score = MaxScore;
                    result.Details.Add("A1 da duoc can trai.");
                }
                else
                {
                    result.Errors.Add($"A1 chua can trai. Hi?n t?i: {ws.Cells["A1"].Style.HorizontalAlignment}.");
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




