using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project12
{
    public class P12T5Grader : ITaskGrader
    {
        public string TaskId => "P12-T5";
        public string TaskName => "Prices: cong thuc Inventory Notice IF <15% thi Low, nguoc lai de trong";
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
                var ws = P12GraderHelpers.GetSheet(studentSheet.Workbook, "Prices");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'Prices'.");
                    return result;
                }

                decimal score = 0m;
                var table = P12GraderHelpers.FindTableByAddress(ws, "A4:L25");
                if (table == null)
                {
                    result.Errors.Add("Không tìm th?y table Prices A4:L25.");
                    return result;
                }

                var noticeColumn = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name, "Inventory Notice", StringComparison.OrdinalIgnoreCase));
                if (noticeColumn == null)
                {
                    result.Errors.Add("Không tìm th?y c?t 'Inventory Notice'.");
                    return result;
                }

                var normalizedFormula = P12GraderHelpers.NormalizeFormula(noticeColumn.CalculatedColumnFormula);
                var expectedFormula = P12GraderHelpers.NormalizeFormula("IF(Table2[[#This Row],[Inventory Level %]]<15%,\"Low\",\"\")");
                if (string.Equals(normalizedFormula, expectedFormula, StringComparison.Ordinal))
                {
                    score += 3m;
                    result.Details.Add("Công th?c Inventory Notice dung chinh xac.");
                }
                else
                {
                    result.Errors.Add($"Công th?c Inventory Notice chua dúng. Hi?n t?i: '{noticeColumn.CalculatedColumnFormula}'.");
                }

                var hasStrictText = normalizedFormula.Contains("\"LOW\",\"\"", StringComparison.Ordinal);
                if (hasStrictText)
                {
                    score += 1m;
                    result.Details.Add("Noi dung text trong IF dung chinh ta: 'Low' va rong.");
                }
                else
                {
                    result.Errors.Add("Text trong cong thuc IF chua dúng ('Low' / \"\").");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"L?i: {ex.Message}");
            }

            return result;
        }
    }
}




