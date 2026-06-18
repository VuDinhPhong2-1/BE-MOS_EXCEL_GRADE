using System.Text.RegularExpressions;
using System.Xml.Linq;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project09
{
    public sealed class WP09T6Grader : IWordTaskGrader
    {
        public string TaskId => "W09-T06";
        public string TaskName => "Insert Full-Length Film citation source";
        public decimal MaxScore => 1m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP09GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);

            var searchableText = BuildCitationSearchText(studentDocument);
            var hasAuthorMetadata = HasSmithJohnAuthor(searchableText);
            var hasTitleMetadata = HasExactPhrase(searchableText, "Event Videography Trends");
            var hasYearMetadata = HasExactPhrase(searchableText, "2019");
            var hasPublisherMetadata = HasExactPhrase(searchableText, "Southridge Press");
            var hasSourceMetadata = hasAuthorMetadata
                || hasTitleMetadata
                || hasYearMetadata
                || hasPublisherMetadata;
            var hasCompleteSourceMetadata = hasAuthorMetadata
                && hasTitleMetadata
                && hasYearMetadata
                && hasPublisherMetadata;

            var fullLengthFilmParagraph = WP09GraderHelpers.GetParagraphs(studentDocument).FirstOrDefault(paragraph =>
                WP09GraderHelpers.GetParagraphText(paragraph).Contains("Full-Length Film", StringComparison.OrdinalIgnoreCase));
            var fullLengthFilmText = WP09GraderHelpers.GetParagraphText(fullLengthFilmParagraph ?? new XElement(WP09GraderHelpers.W + "p"));
            var hasVisibleCitationNearFullLengthFilm = HasSmith2019CitationAfterFullLengthFilm(fullLengthFilmText);
            var hasCitationFieldNearFullLengthFilm = fullLengthFilmParagraph?
                .Descendants(WP09GraderHelpers.W + "instrText")
                .Any(node => node.Value.Contains("CITATION", StringComparison.OrdinalIgnoreCase)) == true;

            if (!hasVisibleCitationNearFullLengthFilm)
            {
                WP09GraderHelpers.AddError(
                    result,
                    "Chưa thấy citation (Smith, 2019) ở bên phải/gần Full-Length Film.",
                    "Đặt con trỏ ở bên phải phần Full-Length Film, vào References > Insert Citation và chèn nguồn John Smith, 2019 để hiển thị dạng (Smith, 2019).");
            }
            else if (hasSourceMetadata && !hasCompleteSourceMetadata)
            {
                WP09GraderHelpers.AddError(
                    result,
                    "Citation hiển thị đúng nhưng thông tin nguồn trong tài liệu chưa khớp Author: Smith, John / Title: Event Videography Trends / Year: 2019 / Publisher: Southridge Press.",
                    "Mở Manage Sources/Edit Source và kiểm tra Author: Smith, John (hoặc John Smith), Title: Event Videography Trends, Year: 2019, Publisher: Southridge Press.");
            }

            if (result.Errors.Count == 0)
            {
                var metadataDetail = hasSourceMetadata
                    ? "metadata nguồn Smith, John / Event Videography Trends / 2019 / Southridge Press được xác minh trong XML."
                    : "không tìm thấy metadata nguồn trong package nên chấm theo citation hiển thị của file.";
                var fieldDetail = hasCitationFieldNearFullLengthFilm ? " Có field CITATION tại đoạn này." : string.Empty;
                result.Details.Add($"Phát hiện citation (Smith, 2019) gần Full-Length Film; {metadataDetail}{fieldDetail}");
            }

            return result;
        }

        private static bool HasSmithJohnAuthor(string searchableText)
        {
            return Regex.IsMatch(searchableText, @"\bSmith\s*,\s*John\b", RegexOptions.IgnoreCase)
                || Regex.IsMatch(searchableText, @"\bJohn\s+Smith\b", RegexOptions.IgnoreCase)
                || (Regex.IsMatch(searchableText, @"\bSmith\b", RegexOptions.IgnoreCase)
                    && Regex.IsMatch(searchableText, @"\bJohn\b", RegexOptions.IgnoreCase));
        }

        private static bool HasExactPhrase(string searchableText, string phrase)
        {
            var escapedPhrase = Regex.Escape(phrase).Replace(@"\ ", @"\s+");
            return Regex.IsMatch(searchableText, $@"(?<![\p{{L}}\p{{N}}]){escapedPhrase}(?![\p{{L}}\p{{N}}])", RegexOptions.IgnoreCase);
        }

        private static bool HasSmith2019CitationAfterFullLengthFilm(string paragraphText)
        {
            var normalized = WP09GraderHelpers.NormalizeText(paragraphText)
                .Replace('\u00A0', ' ')
                .Replace('\u2011', '-');

            var fullLengthFilmIndex = normalized.IndexOf("Full-Length Film", StringComparison.OrdinalIgnoreCase);
            if (fullLengthFilmIndex < 0)
            {
                return false;
            }

            var textAfterHeading = normalized[(fullLengthFilmIndex + "Full-Length Film".Length)..];
            return Regex.IsMatch(textAfterHeading, @"\(\s*Smith\s*,\s*2019\s*\)", RegexOptions.IgnoreCase);
        }

        private static string BuildCitationSearchText(WordGradingContext context)
        {
            var parts = new List<string>();

            foreach (var xmlPart in context.XmlParts)
            {
                if (xmlPart.Key.Contains("bibliography", StringComparison.OrdinalIgnoreCase)
                    || xmlPart.Key.Contains("customXml", StringComparison.OrdinalIgnoreCase)
                    || xmlPart.Key.Contains("document.xml", StringComparison.OrdinalIgnoreCase))
                {
                    parts.Add(xmlPart.Value.ToString(SaveOptions.DisableFormatting));
                }
            }

            return WP09GraderHelpers.NormalizeText(string.Join(" ", parts));
        }
    }
}
