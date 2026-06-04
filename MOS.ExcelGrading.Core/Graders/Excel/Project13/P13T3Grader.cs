using System.Globalization;
using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project13
{
    public class P13T3Grader : ITaskGrader
    {
        public string TaskId => "P13-T3";
        public string TaskName => "Shirt Orders: cong thuc C3 COUNTIF cho Large";
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
                var ws = P13GraderHelpers.GetSheet(studentSheet.Workbook, "Shirt Orders");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Shirt Orders'.");
                    return result;
                }

                var isMatched = P13GraderHelpers.FormulaMatchesAny(
                    ws.Cells["C3"].Formula,
                    "COUNTIF(E:E,\"Large\")",
                    "COUNTIF(E6:E199,\"Large\")");

                if (isMatched)
                {
                    result.Score = MaxScore;
                    result.Details.Add("Công thức C3 dung chinh xac.");
                }
                else
                {
                    result.Errors.Add(
                        $"Công thức C3 chưa đúng. Chap nhan COUNTIF(E:E,\"Large\") hoac COUNTIF(E6:E199,\"Large\"). Hiện tại: '{ws.Cells["C3"].Formula}'.");
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



