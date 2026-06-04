using System.Xml.Linq;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project03
{
    public class WP03T2Grader : IWordTaskGrader
    {
        public string TaskId => "W03-T2";
        public string TaskName => "Trong phần \"Depanning\", chèn biểu tượng nhiệt kế trước cụm từ \"The muffin tray will still be hot!\". Sử dụng phông \"Webdings\" và mã ký tự \"225\".";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };
            const string fixAction = "Trong phần Depanning, đặt con trỏ ngay trước câu \"The muffin tray will still be hot!\", vào Insert > Symbol > More Symbols, chọn font Webdings, nhập mã ký tự 225 rồi chèn biểu tượng nhiệt kế.";

            try
            {
                var bodyElements = WP03GraderHelpers.GetBodyElements(studentDocument);
                var headingIndex = WP03GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Depanning");
                if (headingIndex < 0)
                {
                    WP03GraderHelpers.AddError(result, "Không tìm thấy tiêu đề \"Depanning\".", "Kiểm tra lại tài liệu và đảm bảo vẫn còn tiêu đề \"Depanning\" đúng chính tả trước khi chèn biểu tượng cảnh báo.");
                    return result;
                }

                result.Score += 3m;
                result.Details.Add("Đã tìm thấy đúng phần \"Depanning\".");

                var sectionParagraphs = WP03GraderHelpers.GetSectionParagraphs(bodyElements, headingIndex, stopAtHeading1: false);
                var targetParagraph = WP03GraderHelpers.FindParagraphContainingText(sectionParagraphs, "muffin tray will still be hot");
                if (targetParagraph == null)
                {
                    WP03GraderHelpers.AddError(result, "Không tìm thấy đoạn chứa cụm \"The muffin tray will still be hot!\" trong phần Depanning.", "Khôi phục hoặc nhập lại câu \"The muffin tray will still be hot!\" trong phần Depanning, sau đó chèn biểu tượng nhiệt kế ngay phía trước câu này.");
                    return result;
                }

                const string exactPhrase = "The muffin tray will still be hot!";
                var paragraphText = WP03GraderHelpers.GetParagraphText(targetParagraph);
                if (paragraphText.Contains(exactPhrase, StringComparison.Ordinal))
                {
                    result.Score += 4m;
                    result.Details.Add("Cụm chữ \"The muffin tray will still be hot!\" đúng chính tả và đúng dấu chấm than.");
                }
                else
                {
                    WP03GraderHelpers.AddError(
                        result,
                        $"Cụm chữ cảnh báo chưa đúng tuyệt đối. Giá trị hiện tại là \"{paragraphText}\".",
                        "Sửa cụm cảnh báo thành đúng chính tả và dấu câu: \"The muffin tray will still be hot!\".");
                }

                var runs = targetParagraph.Elements(WP03GraderHelpers.W + "r").ToList();
                var phraseRunIndex = runs.FindIndex(run =>
                {
                    var runText = WP03GraderHelpers.NormalizeText(
                        string.Concat(run.Descendants(WP03GraderHelpers.W + "t").Select(node => node.Value)));
                    return runText.Contains("The muffin tray will still be hot!", StringComparison.Ordinal);
                });

                if (phraseRunIndex < 0)
                {
                    WP03GraderHelpers.AddError(result, "Không xác định được run chứa cụm cảnh báo để kiểm tra vị trí biểu tượng.", fixAction);
                    return result;
                }

                var symbolBeforePhrase = runs.Take(phraseRunIndex).FirstOrDefault(run => run.Element(WP03GraderHelpers.W + "sym") != null);
                if (symbolBeforePhrase == null)
                {
                    WP03GraderHelpers.AddError(result, "Chưa có biểu tượng chèn trước cụm cảnh báo.", fixAction);
                    return result;
                }

                var symbolNode = symbolBeforePhrase.Element(WP03GraderHelpers.W + "sym");
                var symbolFont = symbolNode?.Attribute(WP03GraderHelpers.W + "font")?.Value ?? string.Empty;
                var symbolChar = symbolNode?.Attribute(WP03GraderHelpers.W + "char")?.Value ?? string.Empty;

                var fontCorrect = string.Equals(symbolFont, "Webdings", StringComparison.OrdinalIgnoreCase);
                var charCorrect = string.Equals(symbolChar, "F0E1", StringComparison.OrdinalIgnoreCase)
                                  || string.Equals(symbolChar, "00E1", StringComparison.OrdinalIgnoreCase);

                if (fontCorrect && charCorrect)
                {
                    result.Score += 8m;
                    result.Details.Add("Biểu tượng trước cụm cảnh báo đúng phông Webdings và đúng mã ký tự 225.");
                }
                else
                {
                    WP03GraderHelpers.AddError(
                        result,
                        $"Biểu tượng chưa đúng chuẩn Webdings/225. Giá trị hiện tại: font=\"{symbolFont}\", char=\"{symbolChar}\".",
                        fixAction);
                }

                var phraseRunText = WP03GraderHelpers.NormalizeText(
                    string.Concat(runs[phraseRunIndex].Descendants(WP03GraderHelpers.W + "t").Select(node => node.Value)));
                if (phraseRunText.StartsWith("The muffin tray will still be hot!", StringComparison.Ordinal))
                {
                    result.Score += 3m;
                    result.Details.Add("Không có khoảng trắng/dấu dư trước cụm cảnh báo trong run chứa nội dung.");
                }
                else
                {
                    WP03GraderHelpers.AddError(result, "Cụm cảnh báo có dấu hoặc khoảng trắng dư trước nội dung chính.", "Xóa các khoảng trắng hoặc ký tự thừa giữa biểu tượng nhiệt kế và câu \"The muffin tray will still be hot!\".");
                }

                var duplicateSymbolCount = targetParagraph.Descendants(WP03GraderHelpers.W + "sym").Count();
                if (duplicateSymbolCount > 1)
                {
                    WP03GraderHelpers.AddError(result, $"Có {duplicateSymbolCount} biểu tượng trong cùng đoạn, cần kiểm tra lại để tránh chèn dư.", "Giữ lại một biểu tượng nhiệt kế Webdings mã 225 ngay trước câu cảnh báo và xóa các biểu tượng dư trong cùng đoạn.");
                }
                else
                {
                    result.Score += 2m;
                    result.Details.Add("Không phát hiện biểu tượng dư trong đoạn cảnh báo.");
                }
            }
            catch (Exception ex)
            {
                WP03GraderHelpers.AddError(result, $"Lỗi khi chấm Task 2: {ex.Message}.", "Đóng file Word nếu đang mở, kiểm tra file .docx không bị hỏng rồi tải lại để chấm lại Task 2.");
            }

            return result;
        }
    }
}
