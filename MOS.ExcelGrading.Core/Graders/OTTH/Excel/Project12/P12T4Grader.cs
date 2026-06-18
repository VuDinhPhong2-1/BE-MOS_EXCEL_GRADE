using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project12
{
    public class P12T4Grader : ITaskGrader
    {
        public string TaskId => "P12-T4";
        public string TaskName => "Prices: cong thuc cot Tax dung Unit price * L$2";
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

                var taxColumn = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name, "Tax", StringComparison.OrdinalIgnoreCase));
                if (taxColumn == null)
                {
                    result.Errors.Add("Không tìm th?y c?t 'Tax' trong table Prices.");
                    return result;
                }

                var normalizedFormula = P12GraderHelpers.NormalizeFormula(taxColumn.CalculatedColumnFormula);
                var expectedFormula = P12GraderHelpers.NormalizeFormula("Table2[[#This Row],[Unit price]]*L$2");
                if (string.Equals(normalizedFormula, expectedFormula, StringComparison.Ordinal))
                {
                    score += 3m;
                    result.Details.Add("Công th?c c?t Tax dung chinh xac theo structured reference.");
                }
                else
                {
                    result.Errors.Add($"Công th?c c?t Tax chua dúng. Hi?n t?i: '{taxColumn.CalculatedColumnFormula}'.");
                }

                if (normalizedFormula.Contains("TABLE2[[#THISROW],[UNITPRICE]]", StringComparison.Ordinal)
                    && normalizedFormula.Contains("L2", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Công th?c Tax dung nguon c?t Unit price va o L$2.");
                }
                else
                {
                    result.Errors.Add("Công th?c Tax khong dung nguon d? li?u Unit price va L$2.");
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




