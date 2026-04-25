using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project07
{
    public class P07T6Grader : ITaskGrader
    {
        public string TaskId => "P07-T6";
        public string TaskName => "Transpose Q1 Sales A4:E9 sang Seedling Sales bắt đầu A4";
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
                var source = P07GraderHelpers.GetSheet(studentSheet, "Q1 Sales");
                var target = P07GraderHelpers.GetSheet(studentSheet, "Seedling Sales");
                if (source == null || target == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Q1 Sales' hoặc 'Seedling Sales'.");
                    return result;
                }

                const int sourceStartRow = 4;
                const int sourceStartCol = 1;
                const int sourceRows = 6;   // A4:E9
                const int sourceCols = 5;
                const int targetStartRow = 4;
                const int targetStartCol = 1;
                var total = sourceRows * sourceCols;
                var matched = 0;
                (int SrcRow, int SrcCol, int DestRow, int DestCol, string Expected, string Actual)? firstMismatch = null;

                for (var r = 0; r < sourceRows; r++)
                {
                    for (var c = 0; c < sourceCols; c++)
                    {
                        var srcRow = sourceStartRow + r;
                        var srcCol = sourceStartCol + c;
                        var dstRow = targetStartRow + c;
                        var dstCol = targetStartCol + r;

                        var expected = (source.Cells[srcRow, srcCol].Text ?? string.Empty).Trim();
                        var actual = (target.Cells[dstRow, dstCol].Text ?? string.Empty).Trim();
                        if (string.Equals(expected, actual, StringComparison.Ordinal))
                        {
                            matched++;
                        }
                        else if (firstMismatch == null)
                        {
                            firstMismatch = (srcRow, srcCol, dstRow, dstCol, expected, actual);
                        }
                    }
                }

                var score = Math.Round(MaxScore * matched / total, 2);
                if (matched == total)
                {
                    result.Details.Add("Đã transpose đúng toàn bộ dữ liệu từ Q1 Sales sang Seedling Sales.");
                }
                else
                {
                    result.Errors.Add($"Transpose chưa đúng ({matched}/{total} ô khớp).");
                    if (firstMismatch.HasValue)
                    {
                        var mm = firstMismatch.Value;
                        result.Errors.Add(
                            $"Lệch tại Q1 Sales({mm.SrcRow},{mm.SrcCol}) -> Seedling Sales({mm.DestRow},{mm.DestCol}), mong đợi '{mm.Expected}', hiện tai '{mm.Actual}'.");
                    }
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
