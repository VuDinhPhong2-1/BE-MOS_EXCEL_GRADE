using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project11
{
    public sealed class WP11T5Grader : IWordTaskGrader
    {
        private const int ExpectedColumnSpaceTwips = 454;

        public string TaskId => "W11-T05";
        public string TaskName => "Chia 4 doan van phia tren hinh anh thanh 2 cot, cach nhau 0.8 cm";
        public decimal MaxScore => 25m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP11GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);

            try
            {
                using var document = WP11GraderHelpers.OpenReadOnlyDocument(studentDocument);
                var body = WP11GraderHelpers.GetBody(document);
                var paragraphs = WP11GraderHelpers.GetTopLevelParagraphs(document);
                var imageIndex = WP11GraderHelpers.FindMainContentImageParagraphIndex(paragraphs);

                if (imageIndex < 0)
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Khong xac dinh duoc hinh anh chinh de suy ra 4 doan van can chia cot.",
                        "Bôi đen 4 đoạn văn nằm phía trên hình ảnh (từ 'This preview event...' đến '...even begins.'), vào tab Layout, chọn Columns > More Columns, chọn Presets là Two, tại ô Spacing nhập 0.8 cm (hoặc 0.31\"), tích chọn Equal column width, mục Apply to chọn Selected text, rồi click OK và lưu file.");
                    return result;
                }

                var targetParagraphs = paragraphs
                    .Take(imageIndex)
                    .Where(paragraph => !string.IsNullOrWhiteSpace(WP11GraderHelpers.GetParagraphText(paragraph)))
                    .TakeLast(4)
                    .ToList();

                if (targetParagraphs.Count < 4)
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Khong du 4 doan van ngay phia tren hinh anh de kiem tra chia cot.",
                        "Bôi đen 4 đoạn văn nằm phía trên hình ảnh (từ 'This preview event...' đến '...even begins.'), vào tab Layout, chọn Columns > More Columns, chọn Presets là Two, tại ô Spacing nhập 0.8 cm (hoặc 0.31\"), tích chọn Equal column width, mục Apply to chọn Selected text, rồi click OK và lưu file.");
                    return result;
                }

                var firstText = WP11GraderHelpers.GetParagraphText(targetParagraphs[0]);
                var lastText = WP11GraderHelpers.GetParagraphText(targetParagraphs[^1]);
                if (!firstText.Contains("This preview event", StringComparison.OrdinalIgnoreCase)
                    || !lastText.Contains("trip even begins", StringComparison.OrdinalIgnoreCase))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "4 doan van duoc nhan dien khong dung pham vi yeu cau tu 'This preview event...' den '...even begins.'",
                        "Bôi đen 4 đoạn văn nằm phía trên hình ảnh (từ 'This preview event...' đến '...even begins.'), vào tab Layout, chọn Columns > More Columns, chọn Presets là Two, tại ô Spacing nhập 0.8 cm (hoặc 0.31\"), tích chọn Equal column width, mục Apply to chọn Selected text, rồi click OK và lưu file.");
                    return result;
                }

                var targetSections = targetParagraphs
                    .Select(paragraph => WP11GraderHelpers.GetEffectiveSectionProperties(paragraph, body))
                    .ToList();

                if (targetSections.Any(section => section == null)
                    || targetSections.Select(section => section!.OuterXml).Distinct(StringComparer.Ordinal).Count() != 1)
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "4 doan van muc tieu khong nam trong cung mot section duoc cau hinh cot.",
                        "Bôi đen 4 đoạn văn nằm phía trên hình ảnh (từ 'This preview event...' đến '...even begins.'), vào tab Layout, chọn Columns > More Columns, chọn Presets là Two, tại ô Spacing nhập 0.8 cm (hoặc 0.31\"), tích chọn Equal column width, mục Apply to chọn Selected text, rồi click OK và lưu file.");
                    return result;
                }

                var targetSection = targetSections[0]!;
                if (!WP11GraderHelpers.HasExpectedColumns(targetSection, 2, ExpectedColumnSpaceTwips))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Section cua 4 doan van muc tieu chua co w:cols voi num=2 va space=454 twips (0,8 cm).",
                        "Bôi đen 4 đoạn văn nằm phía trên hình ảnh (từ 'This preview event...' đến '...even begins.'), vào tab Layout, chọn Columns > More Columns, chọn Presets là Two, tại ô Spacing nhập 0.8 cm (hoặc 0.31\"), tích chọn Equal column width, mục Apply to chọn Selected text, rồi click OK và lưu file.");
                }

                var firstTargetIndex = paragraphs.IndexOf(targetParagraphs[0]);
                var previousMeaningfulParagraph = paragraphs
                    .Take(firstTargetIndex)
                    .LastOrDefault(paragraph => !string.IsNullOrWhiteSpace(WP11GraderHelpers.GetParagraphText(paragraph)));

                if (previousMeaningfulParagraph != null)
                {
                    var previousSection = WP11GraderHelpers.GetEffectiveSectionProperties(previousMeaningfulParagraph, body);
                    if (previousSection != null
                        && WP11GraderHelpers.HasExpectedColumns(previousSection, 2, ExpectedColumnSpaceTwips))
                    {
                        WP11GraderHelpers.AddError(
                            result,
                            "Pham vi section 2 cot dang bi mo rong sang doan van truoc nhom 4 doan muc tieu.",
                            "Bôi đen 4 đoạn văn nằm phía trên hình ảnh (từ 'This preview event...' đến '...even begins.'), vào tab Layout, chọn Columns > More Columns, chọn Presets là Two, tại ô Spacing nhập 0.8 cm (hoặc 0.31\"), tích chọn Equal column width, mục Apply to chọn Selected text, rồi click OK và lưu file.");
                    }
                }

                var nextMeaningfulIndex = WP11GraderHelpers.FindNextMeaningfulParagraphIndex(paragraphs, paragraphs.IndexOf(targetParagraphs[^1]));
                if (nextMeaningfulIndex >= 0)
                {
                    var nextSection = WP11GraderHelpers.GetEffectiveSectionProperties(paragraphs[nextMeaningfulIndex], body);
                    if (nextSection != null
                        && WP11GraderHelpers.HasExpectedColumns(nextSection, 2, ExpectedColumnSpaceTwips))
                    {
                        WP11GraderHelpers.AddError(
                            result,
                            "Pham vi section 2 cot dang bi mo rong sang noi dung sau nhom 4 doan muc tieu.",
                            "Bôi đen 4 đoạn văn nằm phía trên hình ảnh (từ 'This preview event...' đến '...even begins.'), vào tab Layout, chọn Columns > More Columns, chọn Presets là Two, tại ô Spacing nhập 0.8 cm (hoặc 0.31\"), tích chọn Equal column width, mục Apply to chọn Selected text, rồi click OK và lưu file.");
                    }
                }

                if (result.Errors.Count == 0)
                {
                    WP11GraderHelpers.AddDetail(result, "4 doan van phia tren hinh anh nam trong cung mot section co w:cols num=2 va w:space=454.");
                }
            }
            catch (Exception ex)
            {
                WP11GraderHelpers.AddError(
                    result,
                    $"Loi khi kiem tra chia cot: {ex.Message}",
                    "Bôi đen 4 đoạn văn nằm phía trên hình ảnh (từ 'This preview event...' đến '...even begins.'), vào tab Layout, chọn Columns > More Columns, chọn Presets là Two, tại ô Spacing nhập 0.8 cm (hoặc 0.31\"), tích chọn Equal column width, mục Apply to chọn Selected text, rồi click OK và lưu file.");
            }

            return result;
        }
    }
}
