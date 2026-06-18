using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project07
{
    public class P07T1Grader : ITaskGrader
    {
        public string TaskId => "P07-T1";
        public string TaskName => "Import Drinks.txt vao sheet Drinks bắt đầu A7 (Use first row as headers)";
        public decimal MaxScore => 4;

        private static readonly (int Row, string A, string B, string C, string D, string E)[] SampleRows =
        {
            (8,  "Smoothie Ancora Artic", "Cocktails", "$4.43", "$4.86", "$5.29"),
            (20, "Chocolate Mudslide", "Cocktails", "$3.78", "$4.32", "$5.40"),
            (40, "Viennese", "Coffee", "$3.24", "$3.78", "$4.32"),
            (60, "Mint Mocha", "Espresso", "$2.80", "$3.23", "$3.88"),
            (86, "Sweetened or Unsweetened Tea", "Tea", "$1.62", "$1.89", "$2.16")
        };

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
                var ws = P07GraderHelpers.GetSheet(studentSheet, "Drinks");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Drinks'.");
                    return result;
                }

                decimal score = 0;
                var headerOk =
                    ws.Cells["A7"].Text == "Drink" &&
                    ws.Cells["B7"].Text == "Type" &&
                    ws.Cells["C7"].Text == "Price for 35 cl" &&
                    ws.Cells["D7"].Text == "Price for 47 cl" &&
                    ws.Cells["E7"].Text == "Price for 59 cl";
                if (headerOk)
                {
                    score += 1m;
                    result.Details.Add("Header tại A7:E7 đúng theo file import.");
                }
                else
                {
                    result.Errors.Add("Header A7:E7 chưa đúng (chưa import với tùy chọn Use first row as headers).");
                }

                var nonEmptyCount = Enumerable.Range(8, 79)
                    .Count(r => !string.IsNullOrWhiteSpace(ws.Cells[r, 1].Text));
                if (nonEmptyCount == 79)
                {
                    score += 1m;
                    result.Details.Add("Đã import đủ 79 dòng dữ liệu (A8:A86).");
                }
                else
                {
                    result.Errors.Add($"Số dòng dữ liệu import chưa đúng: {nonEmptyCount}/79.");
                }

                var sampleMatch = 0;
                foreach (var sample in SampleRows)
                {
                    var ok =
                        ws.Cells[sample.Row, 1].Text == sample.A &&
                        ws.Cells[sample.Row, 2].Text == sample.B &&
                        ws.Cells[sample.Row, 3].Text == sample.C &&
                        ws.Cells[sample.Row, 4].Text == sample.D &&
                        ws.Cells[sample.Row, 5].Text == sample.E;
                    if (ok)
                    {
                        sampleMatch++;
                    }
                }

                score += Math.Round(2m * sampleMatch / SampleRows.Length, 2);
                if (sampleMatch != SampleRows.Length)
                {
                    result.Errors.Add($"Dữ liệu import không khớp mẫu đối chiếu ({sampleMatch}/{SampleRows.Length} dòng mẫu).");
                }
                else
                {
                    result.Details.Add("Dữ liệu import khớp các dòng mẫu quan trọng.");
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

// minor-sync: non-functional graders update

