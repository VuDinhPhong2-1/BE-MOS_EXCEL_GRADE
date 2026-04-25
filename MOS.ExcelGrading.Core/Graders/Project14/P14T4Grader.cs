using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project14
{
    public class P14T4Grader : ITaskGrader
    {
        public string TaskId => "P14-T4";
        public string TaskName => "February: cong thuc Policy Type LEFT(...,2)";
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
                var ws = P14GraderHelpers.GetSheet(studentSheet.Workbook, "February");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'February'.");
                    return result;
                }

                var table = P14GraderHelpers.FindTableByAddress(ws, "A4:G18");
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy table February A4:G18.");
                    return result;
                }

                var column = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name, "Policy Type", StringComparison.OrdinalIgnoreCase));
                if (column == null)
                {
                    result.Errors.Add("Không tìm thấy cột 'Policy Type'.");
                    return result;
                }

                var actual = P14GraderHelpers.NormalizeFormula(column.CalculatedColumnFormula);
                var expected = P14GraderHelpers.NormalizeFormula("LEFT(Table1[[#This Row],[Policy Number]],2)");
                if (string.Equals(actual, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Công thức cột Policy Type dung chinh xac.");
                }
                else
                {
                    result.Errors.Add($"Công thức Policy Type chưa đúng. Hiện tại: '{column.CalculatedColumnFormula}'.");
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



