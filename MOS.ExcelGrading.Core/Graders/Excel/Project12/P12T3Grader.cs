using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.Project12
{
    public class P12T3Grader : ITaskGrader
    {
        public string TaskId => "P12-T3";
        public string TaskName => "Orders: loc theo The House of Alpine Skiing";
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
                var ws = P12GraderHelpers.GetSheet(studentSheet.Workbook, "Orders");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Orders'.");
                    return result;
                }

                decimal score = 0m;
                var table = P12GraderHelpers.FindTableByAddress(ws, "A1:E412");
                if (table != null)
                {
                    score += 1m;
                    result.Details.Add("Tìm thấy table Orders dung range A1:E412.");
                }
                else
                {
                    result.Errors.Add("Không tìm thấy table Orders range A1:E412.");
                    result.Score = score;
                    return result;
                }

                var filteredValue = P12GraderHelpers.GetSingleFilterValue(table, 0);
                if (string.Equals(filteredValue, "The House of Alpine Skiing", StringComparison.Ordinal))
                {
                    score += 2m;
                    result.Details.Add("Dieu kien filter cột dau dung giá trị yêu cầu.");
                }
                else
                {
                    result.Errors.Add($"Gia tri filter chưa đúng. Hiện tại: '{filteredValue}'.");
                }

                if (table.TableStyle == TableStyles.Light18)
                {
                    score += 1m;
                    result.Details.Add("Table style Orders dung: TableStyleLight18.");
                }
                else
                {
                    result.Errors.Add($"Table style Orders chưa đúng. Hiện tại: {table.TableStyle}.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi: {ex.Message}");
            }

            return result;
        }
    }
}



