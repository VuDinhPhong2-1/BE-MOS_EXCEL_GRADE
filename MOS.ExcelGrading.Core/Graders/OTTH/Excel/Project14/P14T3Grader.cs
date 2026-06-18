using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project14
{
    public class P14T3Grader : ITaskGrader
    {
        public string TaskId => "P14-T3";
        public string TaskName => "February: cong thuc Discount";
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
                var ws = P14GraderHelpers.GetSheet(studentSheet.Workbook, "February");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'February'.");
                    return result;
                }

                var table = P14GraderHelpers.FindTableByAddress(ws, "A4:G18");
                if (table == null)
                {
                    result.Errors.Add("Không tìm th?y table February A4:G18.");
                    return result;
                }

                var column = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name, "Discount", StringComparison.OrdinalIgnoreCase));
                if (column == null)
                {
                    result.Errors.Add("Không tìm th?y c?t 'Discount'.");
                    return result;
                }

                var actual = P14GraderHelpers.NormalizeFormula(column.CalculatedColumnFormula);
                var expected = P14GraderHelpers.NormalizeFormula("IF(Table1[[#This Row],[Years as Member]]>3,\"Yes\",\"No\")");
                if (string.Equals(actual, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Công th?c c?t Discount dung chinh xac.");
                }
                else
                {
                    result.Errors.Add($"Công th?c Discount chua dúng. Hi?n t?i: '{column.CalculatedColumnFormula}'.");
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




