using System.Xml;
using System.Text.RegularExpressions;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project15
{
    public class P15T1Grader : ITaskGrader
    {
        public string TaskId => "P15-T1";
        public string TaskName => "Products: dinh dang so Weight voi 3 chu so thap phan";
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
                var ws = P15GraderHelpers.GetSheet(studentSheet.Workbook, "Products");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Products'.");
                    return result;
                }

                var table = ws.Tables.FirstOrDefault(t =>
                    t.Columns.Any(c => string.Equals(c.Name, "Weight", StringComparison.OrdinalIgnoreCase)));
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy table co cột 'Weight'.");
                    return result;
                }

                var weightOffset = table.Columns
                    .Select((c, idx) => new { Column = c, Index = idx })
                    .First(x => string.Equals(x.Column.Name, "Weight", StringComparison.OrdinalIgnoreCase))
                    .Index;
                var weightCol = table.Address.Start.Column + weightOffset;

                var dataStart = table.Address.Start.Row + 1;
                var dataEnd = table.Address.End.Row;
                var totalRows = Math.Max(0, dataEnd - dataStart + 1);
                var validRows = 0;
                for (var row = dataStart; row <= dataEnd; row++)
                {
                    var format = ws.Cells[row, weightCol].Style.Numberformat.Format;
                    if (P15GraderHelpers.IsThreeDecimalNumberFormat(format))
                    {
                        validRows++;
                    }
                }

                if (validRows == totalRows && totalRows > 0)
                {
                    result.Score = MaxScore;
                    result.Details.Add($"Định dạng 3 so thap phan dung tren toàn bộ cột Weight ({validRows}/{totalRows}).");
                }
                else
                {
                    result.Score = 2m;
                    result.Errors.Add($"Định dạng cột Weight chưa đúng tren toàn bộ dữ liệu ({validRows}/{totalRows}).");
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



