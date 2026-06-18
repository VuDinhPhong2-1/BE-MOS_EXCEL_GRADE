using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project12
{
    public class P12T3Grader : ITaskGrader
    {
        public string TaskId => "P12-T3";
        public string TaskName => "Orders: loc theo The House of Alpine Skiing";
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
                var ws = P12GraderHelpers.GetSheet(studentSheet.Workbook, "Orders");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'Orders'.");
                    return result;
                }

                decimal score = 0m;
                var table = P12GraderHelpers.FindTableByAddress(ws, "A1:E412");
                if (table != null)
                {
                    score += 1m;
                    result.Details.Add("Tìm th?y table Orders dung range A1:E412.");
                }
                else
                {
                    result.Errors.Add("Không tìm th?y table Orders range A1:E412.");
                    result.Score = score;
                    return result;
                }

                var filteredValue = P12GraderHelpers.GetSingleFilterValue(table, 0);
                if (string.Equals(filteredValue, "The House of Alpine Skiing", StringComparison.Ordinal))
                {
                    score += 2m;
                    result.Details.Add("Dieu kien filter c?t dau dung giá tr? yêu c?u.");
                }
                else
                {
                    result.Errors.Add($"Gia tri filter chua dúng. Hi?n t?i: '{filteredValue}'.");
                }

                if (table.TableStyle == TableStyles.Light18)
                {
                    score += 1m;
                    result.Details.Add("Table style Orders dung: TableStyleLight18.");
                }
                else
                {
                    result.Errors.Add($"Table style Orders chua dúng. Hi?n t?i: {table.TableStyle}.");
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




