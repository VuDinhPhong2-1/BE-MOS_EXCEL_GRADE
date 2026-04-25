using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project14
{
    public class P14T6Grader : ITaskGrader
    {
        public string TaskId => "P14-T6";
        public string TaskName => "Sales: cong thuc Auction ID RANDBETWEEN(1000,2000)";
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
                var ws = P14GraderHelpers.GetSheet(studentSheet.Workbook, "Sales");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Sales'.");
                    return result;
                }

                var table = P14GraderHelpers.FindTableByAddress(ws, "A3:F17");
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy table Sales A3:F17.");
                    return result;
                }

                var column = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name, "Auction ID", StringComparison.OrdinalIgnoreCase));
                if (column == null)
                {
                    result.Errors.Add("Không tìm thấy cột 'Auction ID'.");
                    return result;
                }

                var actual = P14GraderHelpers.NormalizeFormula(column.CalculatedColumnFormula);
                var expected = P14GraderHelpers.NormalizeFormula("RANDBETWEEN(1000,2000)");
                if (string.Equals(actual, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Công thức Auction ID dung chinh xac.");
                }
                else
                {
                    result.Errors.Add($"Công thức Auction ID chưa đúng. Hiện tại: '{column.CalculatedColumnFormula}'.");
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



