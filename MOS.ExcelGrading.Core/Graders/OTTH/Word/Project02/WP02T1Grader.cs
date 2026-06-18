using System.Xml.Linq;
using MOS.ExcelGrading.Core.Graders.OTTH.Word.Project01;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project02
{
    public class WP02T1Grader : IWordTaskGrader
    {
        private const string FixAction = "Vào tab Review > nhóm Protect > Restrict Editing. Tại mục 2 (Editing restrictions), tích chọn \"Allow only this type of editing in the document\" và chọn \"Tracked changes\". Nhấn \"Yes, Start Enforcing Protection\" và nhập mật khẩu 456.";

        public string TaskId => "W02-T1";
        public string TaskName => "Kích hoạt Track Changes và khóa bằng mật khẩu \"456\".";
        public decimal MaxScore => 12m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                if (!studentDocument.TryGetXmlPart("word/settings.xml", out var settingsXml) || settingsXml.Root == null)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Không tìm thấy word/settings.xml để kiểm tra thiết lập Track Changes.",
                        "Đóng Word, mở lại tệp .docx và lưu lại trước khi chấm; sau đó thực hiện lại thao tác bật Track Changes và khóa chỉnh sửa.");
                    return result;
                }

                var settingsRoot = settingsXml.Root;
                var trackRevisions = settingsRoot.Element(WP01GraderHelpers.W + "trackRevisions");
                if (trackRevisions == null)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Bạn chưa bật Track Changes cho tài liệu.",
                        FixAction);
                }
                else
                {
                    result.Score += 4m;
                    result.Details.Add("Đã tìm thấy thiết lập w:trackRevisions trong word/settings.xml.");
                }

                var documentProtection = settingsRoot.Element(WP01GraderHelpers.W + "documentProtection");
                if (documentProtection == null)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Bạn chưa bật khóa chỉnh sửa Restrict Editing cho chế độ Track Changes.",
                        FixAction);
                    return result;
                }

                result.Score += 2m;
                result.Details.Add("Đã tìm thấy thiết lập w:documentProtection trong word/settings.xml.");

                var editMode = GetWordAttribute(documentProtection, "edit");
                if (!string.Equals(editMode, "trackedChanges", StringComparison.OrdinalIgnoreCase))
                {
                    WP01GraderHelpers.AddError(
                        result,
                        $"Bạn đã chọn sai chế độ khóa chỉnh sửa. Chế độ hiện tại là \"{(string.IsNullOrWhiteSpace(editMode) ? "không có" : editMode)}\".",
                        FixAction);
                }
                else
                {
                    result.Score += 3m;
                    result.Details.Add("Chế độ khóa chỉnh sửa đúng là trackedChanges.");
                }

                var enforcement = GetWordAttribute(documentProtection, "enforcement");
                if (!IsOn(enforcement))
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Bạn đã cấu hình Restrict Editing nhưng chưa Start Enforcing Protection.",
                        FixAction);
                }
                else
                {
                    result.Score += 1m;
                    result.Details.Add("Protection enforcement đã được bật.");
                }

                if (HasPasswordProtection(documentProtection))
                {
                    result.Score += 2m;
                    result.Details.Add("Đã tìm thấy thông tin mã hóa mật khẩu trong w:documentProtection.");
                }
                else
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Bạn chưa đặt mật khẩu bảo vệ Track Changes hoặc file không lưu thông tin hash mật khẩu.",
                        FixAction);
                }
            }
            catch (Exception ex)
            {
                WP01GraderHelpers.AddError(
                    result,
                    $"Lỗi khi chấm Task 1: {ex.Message}.",
                    "Đóng Word, mở lại tệp .docx và lưu lại trước khi chấm; nếu lỗi còn lặp lại, kiểm tra tệp có bị hỏng hay không.");
            }

            return result;
        }

        private static string GetWordAttribute(XElement element, string localName)
        {
            return element.Attribute(WP01GraderHelpers.W + localName)?.Value
                ?? element.Attribute(localName)?.Value
                ?? string.Empty;
        }

        private static bool IsOn(string value)
        {
            return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasPasswordProtection(XElement documentProtection)
        {
            var modernHash = GetWordAttribute(documentProtection, "hash");
            var modernSalt = GetWordAttribute(documentProtection, "salt");
            var legacyHash = GetWordAttribute(documentProtection, "cryptProviderType");
            var legacyAlgorithm = GetWordAttribute(documentProtection, "cryptAlgorithmClass");
            var legacySpinCount = GetWordAttribute(documentProtection, "cryptSpinCount");

            return (!string.IsNullOrWhiteSpace(modernHash) && !string.IsNullOrWhiteSpace(modernSalt))
                || (!string.IsNullOrWhiteSpace(legacyHash)
                    && !string.IsNullOrWhiteSpace(legacyAlgorithm)
                    && !string.IsNullOrWhiteSpace(legacySpinCount));
        }
    }
}
