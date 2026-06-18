using DocumentFormat.OpenXml.Wordprocessing;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project11
{
    public sealed class WP11T6Grader : IWordTaskGrader
    {
        public string TaskId => "W11-T06";
        public string TaskName => "Chèn trang bìa Whisp và cập nhật Title, Subtitle, Author, Company, Date";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP11GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);

            try
            {
                using var document = WP11GraderHelpers.OpenReadOnlyDocument(studentDocument);
                var body = WP11GraderHelpers.GetBody(document);
                var coverBlock = body.Elements<SdtBlock>()
                    .FirstOrDefault(block => block.OuterXml.Contains("Cover Pages", StringComparison.OrdinalIgnoreCase));

                if (coverBlock == null)
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Không tìm thấy cover page ở đầu tài liệu.",
                        "Vào tab Insert, chọn Cover Page > Whisp để chèn trang bìa mẫu Whisp vào đầu tài liệu.");
                    return result;
                }

                if (!WP11GraderHelpers.HasCoverPageAtStart(body, coverBlock))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Cover page không nằm ở đầu tài liệu.",
                        "Vào tab Insert, chọn Cover Page > Whisp để chèn trang bìa mẫu Whisp vào đầu tài liệu.");
                }

                if (!WP11GraderHelpers.HasWhispLikeSignature(coverBlock))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Cover page tìm thấy không có đủ dấu hiệu XML/content control của mẫu Whisp.",
                        "Vào tab Insert, chọn Cover Page > Whisp để chèn trang bìa mẫu Whisp vào đầu tài liệu.");
                }

                var aliasValues = WP11GraderHelpers.GetCoverPageAliasValues(coverBlock);
                aliasValues.TryGetValue("Title", out var titleValues);
                aliasValues.TryGetValue("Subtitle", out var subtitleValues);
                aliasValues.TryGetValue("Author", out var authorValues);
                aliasValues.TryGetValue("Company", out var companyValues);
                aliasValues.TryGetValue("Date", out var dateValues);

                titleValues ??= new List<string>();
                subtitleValues ??= new List<string>();
                authorValues ??= new List<string>();
                companyValues ??= new List<string>();
                dateValues ??= new List<string>();

                if (!WP11GraderHelpers.MatchesExpectedValue(titleValues, "River Cruises Preview"))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Title trên cover page chưa đúng 'River Cruises Preview'.",
                        "Click vào trường [DOCUMENT TITLE] (hoặc Title) trên trang bìa và nhập chính xác 'River Cruises Preview', rồi lưu file.");
                }

                if (!WP11GraderHelpers.MatchesExpectedValue(subtitleValues, "Discover the Best of River Cruising"))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Subtitle trên cover page chưa đúng 'Discover the Best of River Cruising'.",
                        "Click vào trường [DOCUMENT SUBTITLE] (hoặc Subtitle) trên trang bìa và nhập chính xác 'Discover the Best of River Cruising', rồi lưu file.");
                }

                if (!WP11GraderHelpers.MatchesExpectedValue(authorValues, "Margie's Travel", ignoreCase: true)
                    && !WP11GraderHelpers.MatchesExpectedValue(authorValues, "Margie’s Travel", ignoreCase: true))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Author trên cover page chưa đúng 'Margie’s Travel'.",
                        "Click vào trường [AUTHOR] (hoặc Author) trên trang bìa và nhập chính xác 'Margie's Travel' (hoặc 'Margie’s Travel'), rồi lưu file.");
                }

                if (!WP11GraderHelpers.MatchesExpectedValue(companyValues, "MARGIE'S TRAVEL AGENCY", ignoreCase: true)
                    && !WP11GraderHelpers.MatchesExpectedValue(companyValues, "MARGIE’S TRAVEL AGENCY", ignoreCase: true)
                    && !WP11GraderHelpers.MatchesExpectedValue(companyValues, "Margie’s travel agency", ignoreCase: true))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Company trên cover page chưa đúng giá trị yêu cầu của Margie's Travel Agency.",
                        "Click vào trường [COMPANY] (hoặc Company) trên trang bìa và nhập chính xác 'MARGIE'S TRAVEL AGENCY' (hoặc 'MARGIE’S TRAVEL AGENCY'), rồi lưu file.");
                }

                if (!WP11GraderHelpers.HasAcceptableDate(dateValues, coverBlock.OuterXml, DateTime.Today))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Date trên cover page không phải ngày hiện tại và cũng không thể hiện date control hợp lệ.",
                        "Click vào trường [DATE] trên trang bìa, chọn ngày hiện tại (hoặc Today), rồi lưu file.");
                }

                if (!WP11GraderHelpers.ContainsExpectedBodyContent(document))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Sau khi chèn cover page, nội dung chính của tài liệu không còn đầy đủ phía sau.",
                        "Vào tab Insert, chọn Cover Page > Whisp. Cập nhật các trường thông tin: Title là 'River Cruises Preview', Subtitle là 'Discover the Best of River Cruising', Author là 'Margie's Travel', Company là 'MARGIE'S TRAVEL AGENCY', và chọn Date là ngày hiện tại, rồi lưu file.");
                }

                if (result.Errors.Count == 0)
                {
                    WP11GraderHelpers.AddDetail(result, "Trang bìa đầu tài liệu có dấu hiệu Whisp và các content control Title/Subtitle/Author/Company/Date hợp lệ.");
                }
            }
            catch (Exception ex)
            {
                WP11GraderHelpers.AddError(
                    result,
                    $"Lỗi khi kiểm tra cover page: {ex.Message}",
                    "Vào tab Insert, chọn Cover Page > Whisp. Cập nhật các trường thông tin: Title là 'River Cruises Preview', Subtitle là 'Discover the Best of River Cruising', Author là 'Margie's Travel', Company là 'MARGIE'S TRAVEL AGENCY', và chọn Date là ngày hiện tại, rồi lưu file.");
            }

            return result;
        }
    }
}

