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

        public TaskResult Grade(WordGradingContext studentDocument, WordGradingContext? answerDocument = null)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var bodyElements = WP03GraderHelpers.GetBodyElements(studentDocument);
                var headingIndex = WP03GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Depanning");
                if (headingIndex < 0)
                {
                    result.Errors.Add("Không tìm thấy tiêu đề \"Depanning\".");
                    return result;
                }

                result.Score += 3m;
                result.Details.Add("Đã tìm thấy đúng phần \"Depanning\".");

                var sectionParagraphs = WP03GraderHelpers.GetSectionParagraphs(bodyElements, headingIndex, stopAtHeading1: false);
                var targetParagraph = WP03GraderHelpers.FindParagraphContainingText(sectionParagraphs, "muffin tray will still be hot");
                if (targetParagraph == null)
                {
                    result.Errors.Add("Không tìm thấy đoạn chứa cụm \"The muffin tray will still be hot!\" trong phần Depanning.");
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
                    result.Errors.Add(
                        $"Cụm chữ cảnh báo chưa đúng tuyệt đối. Giá trị hiện tại là \"{paragraphText}\".");
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
                    result.Errors.Add("Không xác định được run chứa cụm cảnh báo để kiểm tra vị trí biểu tượng.");
                    return result;
                }

                var symbolBeforePhrase = runs.Take(phraseRunIndex).FirstOrDefault(run => run.Element(WP03GraderHelpers.W + "sym") != null);
                if (symbolBeforePhrase == null)
                {
                    result.Errors.Add("Chưa có biểu tượng chèn trước cụm cảnh báo.");
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
                    result.Errors.Add(
                        $"Biểu tượng chưa đúng chuẩn Webdings/225. Giá trị hiện tại: font=\"{symbolFont}\", char=\"{symbolChar}\".");
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
                    result.Errors.Add("Cụm cảnh báo có dấu hoặc khoảng trắng dư trước nội dung chính.");
                }

                var duplicateSymbolCount = targetParagraph.Descendants(WP03GraderHelpers.W + "sym").Count();
                if (duplicateSymbolCount > 1)
                {
                    result.Errors.Add($"Có {duplicateSymbolCount} biểu tượng trong cùng đoạn, cần kiểm tra lại để tránh chèn dư.");
                }
                else
                {
                    result.Score += 2m;
                    result.Details.Add("Không phát hiện biểu tượng dư trong đoạn cảnh báo.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 2: {ex.Message}.");
            }

            return result;
        }
    }
}
