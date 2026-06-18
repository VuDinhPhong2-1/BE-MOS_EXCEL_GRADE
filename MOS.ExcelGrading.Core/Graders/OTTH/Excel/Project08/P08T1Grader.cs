using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project08
{
    public class P08T1Grader : ITaskGrader
    {
        public string TaskId => "P08-T1";
        public string TaskName => "Summary A2 tạo hyperlink đến www.nodpublishers.com với ScreenTip";
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
                var ws = P08GraderHelpers.GetSheet(studentSheet, "Summary");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Summary'.");
                    return result;
                }

                decimal score = 0;
                var hyperlink = ws.Cells["A2"].Hyperlink;
                if (hyperlink != null)
                {
                    score += 1m;
                    result.Details.Add("Ô A2 đã có hyperlink.");
                }
                else
                {
                    result.Errors.Add("Ô A2 chưa có hyperlink.");
                    result.Score = score;
                    return result;
                }

                var linkText = hyperlink.OriginalString ?? hyperlink.ToString() ?? string.Empty;
                var normalized = linkText.TrimEnd('/').ToLowerInvariant();
                if (normalized == "http://www.nodpublishers.com")
                {
                    score += 2m;
                    result.Details.Add("Hyperlink A2 đúng URL yêu cầu.");
                }
                else
                {
                    result.Errors.Add($"URL hyperlink A2 chưa đúng. Hiện tại: '{linkText}'.");
                }

                var ns = new XmlNamespaceManager(ws.WorksheetXml.NameTable);
                ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var linkNode = ws.WorksheetXml.SelectSingleNode("//x:hyperlinks/x:hyperlink[@ref='A2']", ns);
                var tooltip = linkNode?.Attributes?["tooltip"]?.Value ?? string.Empty;
                var displayText = ws.Cells["A2"].Text ?? string.Empty;
                var tipOrDisplayOk =
                    string.Equals(tooltip, "Company Website", StringComparison.Ordinal) ||
                    string.Equals(displayText, "nodpublishers.com", StringComparison.OrdinalIgnoreCase);
                if (tipOrDisplayOk)
                {
                    score += 1m;
                    result.Details.Add("Thông tin hiển thị hyperlink A2 đúng yêu cầu.");
                }
                else
                {
                    result.Errors.Add($"ScreenTip/Display text chưa đúng. Tooltip='{tooltip}', Text='{displayText}'.");
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

