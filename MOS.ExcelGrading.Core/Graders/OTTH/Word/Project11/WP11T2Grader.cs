using DocumentFormat.OpenXml.Wordprocessing;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project11
{
    public sealed class WP11T2Grader : IWordTaskGrader
    {
        public string TaskId => "W11-T02";
        public string TaskName => "Xóa header, footer và watermark nhưng giữ nguyên các thông tin khác";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP11GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);

            try
            {
                using var document = WP11GraderHelpers.OpenReadOnlyDocument(studentDocument);
                var sections = WP11GraderHelpers.GetSectionProperties(document);

                if (sections.Any(section => section.Elements<HeaderReference>().Any()))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Tài liệu vẫn còn header reference trong w:sectPr.",
                        "Vào File > Info, click vào Check for Issues và chọn Inspect Document. Trong hộp thoại Document Inspector, chỉ chọn (tích) ô Headers, Footers, and Watermarks, click Inspect, sau đó click Remove All tại phần Headers, Footers, and Watermarks, đóng hộp thoại rồi lưu file.");
                }

                if (sections.Any(section => section.Elements<FooterReference>().Any()))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Tài liệu vẫn còn footer reference trong w:sectPr.",
                        "Vào File > Info, click vào Check for Issues và chọn Inspect Document. Trong hộp thoại Document Inspector, chỉ chọn (tích) ô Headers, Footers, and Watermarks, click Inspect, sau đó click Remove All tại phần Headers, Footers, and Watermarks, đóng hộp thoại rồi lưu file.");
                }

                if (document.MainDocumentPart?.HeaderParts.Any(part => WP11GraderHelpers.HasVisibleContent(part.Header) || WP11GraderHelpers.HasWatermark(part.Header)) == true)
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Vẫn còn HeaderPart có nội dung hoặc watermark.",
                        "Vào File > Info, click vào Check for Issues và chọn Inspect Document. Trong hộp thoại Document Inspector, chỉ chọn (tích) ô Headers, Footers, and Watermarks, click Inspect, sau đó click Remove All tại phần Headers, Footers, and Watermarks, đóng hộp thoại rồi lưu file.");
                }

                if (document.MainDocumentPart?.FooterParts.Any(part => WP11GraderHelpers.HasVisibleContent(part.Footer)) == true)
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Vẫn còn FooterPart có nội dung.",
                        "Vào File > Info, click vào Check for Issues và chọn Inspect Document. Trong hộp thoại Document Inspector, chỉ chọn (tích) ô Headers, Footers, and Watermarks, click Inspect, sau đó click Remove All tại phần Headers, Footers, and Watermarks, đóng hộp thoại rồi lưu file.");
                }

                if (document.MainDocumentPart != null
                    && (WP11GraderHelpers.HasWatermark(document.MainDocumentPart.Document)
                        || document.MainDocumentPart.HeaderParts.Any(part => WP11GraderHelpers.HasWatermark(part.Header))
                        || document.MainDocumentPart.FooterParts.Any(part => WP11GraderHelpers.HasWatermark(part.Footer))))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Vẫn phát hiện watermark dạng VML/shape trong tài liệu.",
                        "Vào File > Info, click vào Check for Issues và chọn Inspect Document. Trong hộp thoại Document Inspector, chỉ chọn (tích) ô Headers, Footers, and Watermarks, click Inspect, sau đó click Remove All tại phần Headers, Footers, and Watermarks, đóng hộp thoại rồi lưu file.");
                }

                var customXmlCount = studentDocument.Entries.Count(entry =>
                    entry.StartsWith("customXml/item", StringComparison.OrdinalIgnoreCase)
                    && entry.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                    && !entry.Contains("itemProps", StringComparison.OrdinalIgnoreCase)
                    && !entry.Contains("/_rels/", StringComparison.OrdinalIgnoreCase));

                if (customXmlCount == 0)
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Không còn Custom XML Data sau khi Inspect Document.",
                        "Vào File > Info, click vào Check for Issues và chọn Inspect Document. Khi thực hiện, lưu ý KHÔNG click Remove All tại Custom XML Data hoặc Document Properties, chỉ click Remove All ở phần Headers, Footers, and Watermarks, đóng hộp thoại rồi lưu file.");
                }

                if (document.CoreFilePropertiesPart == null
                    || document.ExtendedFilePropertiesPart == null
                    || document.CustomFilePropertiesPart == null)
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Một hoặc nhiều Document Properties đã bị xóa nhầm.",
                        "Vào File > Info, click vào Check for Issues và chọn Inspect Document. Khi thực hiện, lưu ý KHÔNG click Remove All tại Custom XML Data hoặc Document Properties, chỉ click Remove All ở phần Headers, Footers, and Watermarks, đóng hộp thoại rồi lưu file.");
                }

                if (!WP11GraderHelpers.ContainsExpectedBodyContent(document))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Nội dung chính của tài liệu có dấu hiệu bị xóa hoặc thay đổi quá mức trong quá trình Inspect Document.",
                        "Vào File > Info, click vào Check for Issues và chọn Inspect Document. Khi thực hiện, lưu ý KHÔNG click Remove All tại Custom XML Data hoặc Document Properties, chỉ click Remove All ở phần Headers, Footers, and Watermarks, đóng hộp thoại rồi lưu file.");
                }

                if (result.Errors.Count == 0)
                {
                    WP11GraderHelpers.AddDetail(result, "Không còn header/footer/watermark, đồng thời vẫn giữ lại Custom XML, Document Properties và nội dung chính.");
                }
            }
            catch (Exception ex)
            {
                WP11GraderHelpers.AddError(
                    result,
                    $"Lỗi khi kiểm tra việc xóa header/footer/watermark: {ex.Message}",
                    "Vào File > Info, click vào Check for Issues và chọn Inspect Document. Trong hộp thoại Document Inspector, chỉ chọn (tích) ô Headers, Footers, and Watermarks, click Inspect, sau đó click Remove All tại phần Headers, Footers, and Watermarks, đóng hộp thoại rồi lưu file.");
            }

            return result;
        }
    }
}

