using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project11
{
    public class P11T6Grader : ITaskGrader
    {
        public string TaskId => "P11-T6";
        public string TaskName => "Costs: table A3:E26 with calculated columns";
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
                var ws = P11GraderHelpers.GetSheet(studentSheet.Workbook, "Costs");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Costs'.");
                    return result;
                }

                decimal score = 0m;
                var table = P11GraderHelpers.FindTableByAddress(ws, "A3:E26");
                if (table != null)
                {
                    score += 1m;
                    result.Details.Add("Tim thay table dung range A3:E26.");
                }
                else
                {
                    result.Errors.Add($"Khong tim thay table A3:E26. Hien tai: {P11GraderHelpers.JoinTableAddresses(ws)}");
                    result.Score = score;
                    return result;
                }

                if (table.TableStyle == TableStyles.Medium3)
                {
                    score += 1m;
                    result.Details.Add("Table style dung: TableStyleMedium3.");
                }
                else
                {
                    result.Errors.Add($"Table style chua dung. Hien tai: {table.TableStyle}.");
                }

                var markupColumn = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name, "Markup", StringComparison.OrdinalIgnoreCase));
                var marginColumn = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name, "Margin", StringComparison.OrdinalIgnoreCase));

                var markupFormulaOk = markupColumn != null
                                      && string.Equals(
                                          P11GraderHelpers.NormalizeFormula(markupColumn.CalculatedColumnFormula),
                                          P11GraderHelpers.NormalizeFormula("E4/B4"),
                                          StringComparison.Ordinal);
                var marginFormulaOk = marginColumn != null
                                      && string.Equals(
                                          P11GraderHelpers.NormalizeFormula(marginColumn.CalculatedColumnFormula),
                                          P11GraderHelpers.NormalizeFormula("D4-B4"),
                                          StringComparison.Ordinal);

                if (markupFormulaOk)
                {
                    score += 1m;
                    result.Details.Add("Calculated formula cot Markup dung (E4/B4).");
                }
                else
                {
                    result.Errors.Add(
                        $"Formula cot Markup chua dung. Hien tai: '{markupColumn?.CalculatedColumnFormula}'.");
                }

                if (marginFormulaOk)
                {
                    score += 1m;
                    result.Details.Add("Calculated formula cot Margin dung (D4-B4).");
                }
                else
                {
                    result.Errors.Add(
                        $"Formula cot Margin chua dung. Hien tai: '{marginColumn?.CalculatedColumnFormula}'.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }
}
