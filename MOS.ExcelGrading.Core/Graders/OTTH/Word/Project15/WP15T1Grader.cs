using System.Xml.Linq;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project15
{
    public class WP15T1Grader : IWordTaskGrader
    {
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";

        public string TaskId => "W15-T1";
        public string TaskName => "Lưu tài liệu thành mẫu Word 2019 tên Notes, không hỗ trợ macro";
        public decimal MaxScore => 12m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            var sourceFileName = studentDocument.SourceFileName ?? string.Empty;
            var extension = Path.GetExtension(sourceFileName);
            if (!string.Equals(extension, ".dotx", StringComparison.OrdinalIgnoreCase))
            {
                result.Errors.Add("File phải được lưu dưới dạng Word Template không hỗ trợ macro (.dotx).");
                result.FixActions.Add("Mở tài liệu trong Word, chọn File > Save As/Browse, chọn Save as type là Word Template (*.dotx), đặt tên Notes và lưu vào vị trí Templates mặc định.");
                return result;
            }

            result.Score += 3m;
            result.Details.Add("Định dạng file là Word Template không hỗ trợ macro (.dotx).");

            var baseName = Path.GetFileNameWithoutExtension(sourceFileName) ?? string.Empty;
            if (!string.Equals(baseName, "Notes", StringComparison.OrdinalIgnoreCase))
            {
                result.Errors.Add("Tên file mẫu phải là Notes.dotx.");
                result.FixActions.Add("Đổi tên/lưu lại mẫu thành Notes.dotx. Tên file phải là Notes và phần mở rộng phải là .dotx.");
                return result;
            }

            result.Score += 2m;
            result.Details.Add("Tên file đúng yêu cầu Notes.dotx.");

            if (studentDocument.Entries.Any(entry =>
                    entry.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase)
                    || entry.EndsWith(".bin", StringComparison.OrdinalIgnoreCase)
                        && entry.Contains("vba", StringComparison.OrdinalIgnoreCase)))
            {
                result.Errors.Add("Mẫu không được hỗ trợ macro hoặc chứa VBA project.");
                result.FixActions.Add("Lưu lại file bằng loại Word Template (*.dotx), không chọn Word Macro-Enabled Template (*.dotm), và không nhúng macro/VBA vào mẫu.");
                return result;
            }

            result.Score += 2m;
            result.Details.Add("Không phát hiện thành phần macro/VBA trong gói Word.");

            if (!HasTemplateContentType(studentDocument))
            {
                result.Errors.Add("Gói OpenXML chưa thể hiện đúng kiểu tài liệu Word Template (.dotx).");
                result.FixActions.Add("Trong Word, dùng File > Save As và chọn đúng Save as type: Word Template (*.dotx) để Word tạo lại content type của mẫu.");
            }
            else
            {
                result.Score += 2m;
                result.Details.Add("Content type của document part là Word template.");
            }

            if (!IsCompatibleWithCurrentWord(studentDocument))
            {
                result.Errors.Add("Tài liệu chưa được đặt tương thích với các tính năng Word mới nhất.");
                result.FixActions.Add("Mở tài liệu trong Word 2019, nếu có Compatibility Mode hãy dùng File > Info > Convert, sau đó lưu lại dưới dạng Word Template (*.dotx).");
            }
            else
            {
                result.Score += 2m;
                result.Details.Add("Thiết lập compatibilityMode phù hợp Word 2013/2016/2019 hoặc mới hơn.");
            }

            if (!studentDocument.HasMainDocumentPart || studentDocument.MainDocumentXml == null)
            {
                result.Errors.Add("Không đọc được nội dung tài liệu chính trong mẫu.");
                result.FixActions.Add("Mở Notes.dotx trong Word, kiểm tra file không bị lỗi, lưu lại mẫu rồi nộp lại.");
            }
            else
            {
                result.Score += 1m;
                result.Details.Add("Mẫu vẫn chứa phần tài liệu chính hợp lệ.");
            }

            return result;
        }

        private static bool HasTemplateContentType(WordGradingContext context)
        {
            if (!context.TryGetXmlPart("[Content_Types].xml", out var contentTypesXml))
            {
                return false;
            }

            var documentOverride = contentTypesXml.Root?
                .Elements(Ct + "Override")
                .FirstOrDefault(element =>
                    string.Equals(
                        element.Attribute("PartName")?.Value,
                        "/word/document.xml",
                        StringComparison.OrdinalIgnoreCase));

            var contentType = documentOverride?.Attribute("ContentType")?.Value ?? string.Empty;
            return contentType.Contains("wordprocessingml.template.main+xml", StringComparison.OrdinalIgnoreCase)
                && !contentType.Contains("macroEnabled", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsCompatibleWithCurrentWord(WordGradingContext context)
        {
            if (!context.TryGetXmlPart("word/settings.xml", out var settingsXml))
            {
                return true;
            }

            var compatibilityMode = settingsXml.Root?
                .Descendants(W + "compatSetting")
                .FirstOrDefault(element =>
                    string.Equals(element.Attribute(W + "name")?.Value, "compatibilityMode", StringComparison.OrdinalIgnoreCase))
                ?.Attribute(W + "val")
                ?.Value;

            if (string.IsNullOrWhiteSpace(compatibilityMode))
            {
                return true;
            }

            return int.TryParse(compatibilityMode, out var mode) && mode >= 15;
        }
    }
}
