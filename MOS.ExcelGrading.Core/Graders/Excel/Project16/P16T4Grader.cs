using System.Xml;
using System.Text.RegularExpressions;
using System.Globalization;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.Project16
{
    public class P16T4Grader : ITaskGrader
    {
        public string TaskId => "P16-T4";
        public string TaskName => "Products: Table style Medium1";
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
                    result.Errors.Add("Không tìm thấy sheet 'Products'.");
                    return result;
                }

                var table = ws.Tables.FirstOrDefault(t => P16GraderHelpers.IsRangeMatch(t.Address.Address, "A2:G54"));
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy table A2:G54.");
                    return result;
                }

                if (table.TableStyle == TableStyles.Medium1)
                {
                    result.Score = MaxScore;
                    result.Details.Add("Table style dung: TableStyleMedium1.");
                }
                else
                {
                    result.Errors.Add($"Table style chưa đúng. Hiện tại: {table.TableStyle}.");
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



