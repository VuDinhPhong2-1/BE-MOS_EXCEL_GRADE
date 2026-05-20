using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project22
{
    public class P22T3Grader : ITaskGrader
    {
        public string TaskId => "P22-T3";
        public string TaskName => "Trên trang tính \"Task\", định cấu hình các tùy chọn kiểu bảng để mỗi hàng được tự động tô màu xen kẽ.";
        public decimal MaxScore => 16m;

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
                var worksheet = P22GraderHelpers.GetSheet(studentSheet.Workbook, "Task");
                if (worksheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Task'.");
                    return result;
                }

                decimal score = 0m;

                var table = P22GraderHelpers.FindTable(
                    worksheet,
                    "Task",
                    "ID",
                    "Name",
                    "Task 1",
                    "Task 10",
                    "Total Tasks")
                    ?? P22GraderHelpers.FindTable(
                        worksheet,
                        "Table1",
                        "ID",
                        "Name",
                        "Task 1",
                        "Task 10",
                        "Total Tasks");

                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy bảng dữ liệu cần chấm trên sheet Task.");
                    return result;
                }

                score += 4m;
                result.Details.Add("Đã tìm thấy bảng dữ liệu trên sheet Task để kiểm tra kiểu bảng.");

                var styleName = P22GraderHelpers.NormalizeText(table.TableStyle.ToString());
                if (!string.IsNullOrWhiteSpace(styleName) && !string.Equals(styleName, "None", StringComparison.OrdinalIgnoreCase))
                {
                    score += 4m;
                    result.Details.Add($"Bảng đang sử dụng table style '{styleName}'.");
                }
                else
                {
                    result.Errors.Add("Bảng chưa được áp dụng table style hợp lệ.");
                }

                if (table.ShowRowStripes)
                {
                    score += 6m;
                    result.Details.Add("Tùy chọn Banded Rows đã được bật trên bảng.");
                }
                else
                {
                    result.Errors.Add("Tùy chọn Banded Rows chưa được bật trên bảng.");
                }

                var tableXml = table.TableXml;
                var ns = new XmlNamespaceManager(tableXml.NameTable);
                ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var styleInfoNode = tableXml.SelectSingleNode("/x:table/x:tableStyleInfo", ns);
                var showRowStripesAttr = styleInfoNode?.Attributes?["showRowStripes"]?.Value ?? string.Empty;
                if (string.Equals(showRowStripesAttr, "1", StringComparison.OrdinalIgnoreCase))
                {
                    score += 2m;
                    result.Details.Add("Table XML ghi nhận showRowStripes='1' đúng với yêu cầu.");
                }
                else
                {
                    result.Errors.Add(
                        $"Table XML chưa đúng tùy chọn Banded Rows. Giá trị showRowStripes hiện tại: '{showRowStripesAttr}'.");
                }

                if (!table.ShowRowStripes)
                {
                    result.Errors.Add("Lưu ý: Tô màu thủ công từng dòng không được tính là bật Banded Rows.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 3: {ex.Message}.");
            }

            return result;
        }
    }
}
