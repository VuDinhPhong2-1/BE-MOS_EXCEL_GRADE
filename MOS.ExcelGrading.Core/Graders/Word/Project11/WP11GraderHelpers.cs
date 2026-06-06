using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project11
{
    internal static class WP11GraderHelpers
    {
        private static readonly Regex WhiteSpaceRegex = new(@"\s+", RegexOptions.Compiled);
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static TaskResult CreateResult(string taskId, string taskName, decimal maxScore)
        {
            return new TaskResult
            {
                TaskId = taskId,
                TaskName = taskName,
                MaxScore = maxScore,
                Score = maxScore
            };
        }

        public static void AddError(TaskResult result, string errorMessage, string fixAction)
        {
            result.Score = 0m;
            result.Errors.Add(errorMessage);
            // Only add a single fixAction guidance per TaskResult to avoid
            // producing multiple, potentially duplicated suggestions.
            if (!string.IsNullOrWhiteSpace(fixAction)
                && result.FixActions.Count == 0)
            {
                result.FixActions.Add(fixAction);
            }
        }

        public static void AddDetail(TaskResult result, string detail)
        {
            if (!string.IsNullOrWhiteSpace(detail))
            {
                result.Details.Add(detail);
            }
        }

        public static WordprocessingDocument OpenReadOnlyDocument(WordGradingContext context)
        {
            if (context.PackageBytes.Length == 0)
            {
                throw new InvalidOperationException("Khong co du lieu .docx de mo bang WordprocessingDocument.");
            }

            return WordprocessingDocument.Open(new MemoryStream(context.PackageBytes, writable: false), false);
        }

        public static string NormalizeText(string? value)
        {
            return WhiteSpaceRegex.Replace((value ?? string.Empty).Trim(), " ");
        }

        public static string NormalizeComparisonText(string? value)
        {
            return NormalizeText(value)
                .Replace('’', '\'')
                .ToUpperInvariant();
        }

        public static Body GetBody(WordprocessingDocument document)
        {
            return document.MainDocumentPart?.Document?.Body
                ?? throw new InvalidOperationException("Tai lieu thieu MainDocumentPart hoac Body.");
        }

        public static List<Paragraph> GetTopLevelParagraphs(WordprocessingDocument document)
        {
            return GetBody(document).Elements<Paragraph>().ToList();
        }

        public static List<Paragraph> GetMeaningfulParagraphs(WordprocessingDocument document)
        {
            return GetTopLevelParagraphs(document)
                .Where(IsMeaningfulParagraph)
                .ToList();
        }

        public static bool IsMeaningfulParagraph(Paragraph paragraph)
        {
            return !string.IsNullOrWhiteSpace(GetParagraphText(paragraph)) || HasImage(paragraph);
        }

        public static string GetParagraphText(Paragraph paragraph)
        {
            var text = string.Concat(paragraph.Descendants<Text>().Select(node => node.Text));
            return NormalizeText(text);
        }

        public static bool HasImage(Paragraph paragraph)
        {
            return paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any()
                || paragraph.OuterXml.Contains("<w:pict", StringComparison.OrdinalIgnoreCase);
        }

        public static int FindMainContentImageParagraphIndex(IReadOnlyList<Paragraph> paragraphs)
        {
            for (var index = 0; index < paragraphs.Count; index++)
            {
                if (!HasImage(paragraphs[index]))
                {
                    continue;
                }

                var previousNonEmpty = paragraphs
                    .Take(index)
                    .Where(IsMeaningfulParagraph)
                    .TakeLast(4)
                    .Select(GetParagraphText)
                    .ToList();

                var nextNonEmpty = paragraphs
                    .Skip(index + 1)
                    .Where(IsMeaningfulParagraph)
                    .Take(3)
                    .Select(GetParagraphText)
                    .ToList();

                if (previousNonEmpty.Count < 4 || nextNonEmpty.Count == 0)
                {
                    continue;
                }

                var previousText = NormalizeText(string.Join(" ", previousNonEmpty));
                var nextText = NormalizeText(string.Join(" ", nextNonEmpty));

                var matchesContext =
                    previousText.Contains("This preview event will be hosted", StringComparison.OrdinalIgnoreCase)
                    && previousText.Contains("She will give you practical advice", StringComparison.OrdinalIgnoreCase)
                    && previousText.Contains("The event is aimed at people", StringComparison.OrdinalIgnoreCase)
                    && previousText.Contains("We hope to have you join us", StringComparison.OrdinalIgnoreCase)
                    && nextText.Contains("This event will take place", StringComparison.OrdinalIgnoreCase);

                if (matchesContext)
                {
                    return index;
                }
            }

            return -1;
        }

        public static int FindNextMeaningfulParagraphIndex(IReadOnlyList<Paragraph> paragraphs, int startIndex)
        {
            for (var index = Math.Max(0, startIndex + 1); index < paragraphs.Count; index++)
            {
                if (IsMeaningfulParagraph(paragraphs[index]))
                {
                    return index;
                }
            }

            return -1;
        }

        public static bool HasExactLineSpacing(Paragraph paragraph, int expectedTwips)
        {
            var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
            if (spacing == null)
            {
                return false;
            }

            var lineRule = spacing.GetAttribute("lineRule", W.NamespaceName).Value ?? string.Empty;
            var lineValue = spacing.GetAttribute("line", W.NamespaceName).Value;

            return string.Equals(lineRule, "exact", StringComparison.OrdinalIgnoreCase)
                && string.Equals(lineValue, expectedTwips.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);
        }

        public static bool HasStyle(Paragraph paragraph, string expectedStyleId)
        {
            var expected = NormalizeComparisonText(expectedStyleId);
            var paragraphStyle = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (NormalizeComparisonText(paragraphStyle) == expected)
            {
                return true;
            }

            var paragraphMarkStyle = paragraph.ParagraphProperties?
                .ParagraphMarkRunProperties?
                .Descendants<RunStyle>()
                .FirstOrDefault()?
                .Val?
                .Value;
            if (NormalizeComparisonText(paragraphMarkStyle) == expected)
            {
                return true;
            }

            return paragraph.Descendants<RunStyle>()
                .Any(style => NormalizeComparisonText(style.Val?.Value) == expected);
        }

        public static List<SectionProperties> GetSectionProperties(WordprocessingDocument document)
        {
            var body = GetBody(document);
            var sections = new List<SectionProperties>();

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                var sectionProperties = paragraph.ParagraphProperties?.SectionProperties;
                if (sectionProperties != null)
                {
                    sections.Add(sectionProperties);
                }
            }

            var bodySectionProperties = body.Elements<SectionProperties>().FirstOrDefault();
            if (bodySectionProperties != null)
            {
                sections.Add(bodySectionProperties);
            }

            return sections;
        }

        public static SectionProperties? GetEffectiveSectionProperties(Paragraph paragraph, Body body)
        {
            var ownSectionProperties = paragraph.ParagraphProperties?.SectionProperties;
            if (ownSectionProperties != null)
            {
                return ownSectionProperties;
            }

            foreach (var nextParagraph in paragraph.ElementsAfter().OfType<Paragraph>())
            {
                var nextSectionProperties = nextParagraph.ParagraphProperties?.SectionProperties;
                if (nextSectionProperties != null)
                {
                    return nextSectionProperties;
                }
            }

            return body.Elements<SectionProperties>().FirstOrDefault();
        }

        public static bool SectionHasExpectedPageBorders(SectionProperties sectionProperties)
        {
            var pageBorders = sectionProperties.GetFirstChild<PageBorders>();
            if (pageBorders == null)
            {
                return false;
            }

            return HasExpectedBorder(pageBorders.TopBorder)
                && HasExpectedBorder(pageBorders.BottomBorder)
                && HasExpectedBorder(pageBorders.LeftBorder)
                && HasExpectedBorder(pageBorders.RightBorder);
        }

        public static bool HasVisibleContent(OpenXmlPartRootElement? rootElement)
        {
            if (rootElement == null)
            {
                return false;
            }

            return !string.IsNullOrWhiteSpace(NormalizeText(rootElement.InnerText))
                || rootElement.OuterXml.Contains("<w:drawing", StringComparison.OrdinalIgnoreCase)
                || rootElement.OuterXml.Contains("<w:pict", StringComparison.OrdinalIgnoreCase);
        }

        public static bool HasWatermark(OpenXmlPartRootElement? rootElement)
        {
            if (rootElement == null)
            {
                return false;
            }

            var xml = rootElement.OuterXml;
            return xml.Contains("PowerPlusWaterMarkObject", StringComparison.OrdinalIgnoreCase)
                || xml.Contains("watermark", StringComparison.OrdinalIgnoreCase);
        }

        public static Dictionary<string, List<string>> GetCoverPageAliasValues(SdtBlock coverBlock)
        {
            var values = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var coverXml = XDocument.Parse(coverBlock.OuterXml);

            foreach (var sdt in coverXml.Descendants(W + "sdt"))
            {
                var alias = sdt.Element(W + "sdtPr")?
                    .Element(W + "alias")?
                    .Attribute(W + "val")?
                    .Value;

                if (string.IsNullOrWhiteSpace(alias))
                {
                    continue;
                }

                var text = NormalizeText(string.Concat(sdt.Descendants(W + "t").Select(node => node.Value)));
                if (!values.TryGetValue(alias, out var aliasValues))
                {
                    aliasValues = new List<string>();
                    values[alias] = aliasValues;
                }

                if (!string.IsNullOrWhiteSpace(text))
                {
                    aliasValues.Add(text);
                }
            }

            return values;
        }

        public static bool HasWhispLikeSignature(SdtBlock coverBlock)
        {
            var xml = coverBlock.OuterXml;
            return xml.Contains("Cover Pages", StringComparison.OrdinalIgnoreCase)
                && xml.Contains("Title", StringComparison.OrdinalIgnoreCase)
                && xml.Contains("Subtitle", StringComparison.OrdinalIgnoreCase)
                && xml.Contains("Author", StringComparison.OrdinalIgnoreCase)
                && xml.Contains("Company", StringComparison.OrdinalIgnoreCase)
                && xml.Contains("Date", StringComparison.OrdinalIgnoreCase);
        }

        public static bool HasCoverPageAtStart(Body body, SdtBlock coverBlock)
        {
            var leadingElements = body.ChildElements
                .Take(2)
                .ToList();

            return leadingElements.Contains(coverBlock);
        }

        public static bool HasExpectedColumns(SectionProperties sectionProperties, int expectedCount, int expectedSpaceTwips)
        {
            var columns = sectionProperties.GetFirstChild<Columns>();
            if (columns == null)
            {
                return false;
            }

            var actualCount = columns.ColumnCount?.Value ?? 1;
            var actualSpace = columns.Space?.Value;

            return actualCount == expectedCount
                && actualSpace == expectedSpaceTwips.ToString(CultureInfo.InvariantCulture);
        }

        public static bool ContainsExpectedBodyContent(WordprocessingDocument document)
        {
            var normalizedText = NormalizeText(GetBody(document).InnerText);
            return normalizedText.Contains("This preview event will be hosted", StringComparison.OrdinalIgnoreCase)
                && normalizedText.Contains("margie@margiestravel.com", StringComparison.OrdinalIgnoreCase);
        }

        public static bool MatchesExpectedValue(IEnumerable<string> actualValues, string expectedValue, bool ignoreCase = false)
        {
            var comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            var normalizedExpected = ignoreCase
                ? NormalizeComparisonText(expectedValue)
                : NormalizeText(expectedValue);

            if (ignoreCase)
            {
                normalizedExpected = normalizedExpected
                    .Replace("\u2019", "'")
                    .Replace("\u2018", "'");
            }

            return actualValues.Any(value =>
            {
                var normalizedActual = ignoreCase
                    ? NormalizeComparisonText(value)
                    : NormalizeText(value);

                if (ignoreCase)
                {
                    normalizedActual = normalizedActual
                        .Replace("\u2019", "'")
                        .Replace("\u2018", "'");
                }

                return string.Equals(normalizedActual, normalizedExpected, comparison);
            });
        }

        public static bool HasAcceptableDate(IEnumerable<string> dateValues, string coverXml, DateTime today)
        {
            foreach (var dateValue in dateValues)
            {
                if (ContainsToday(dateValue, today))
                {
                    return true;
                }

                if (TryParseSupportedDate(dateValue, out _))
                {
                    return true;
                }
            }

            return coverXml.Contains("/ns0:CoverPageProperties[1]/ns0:PublishDate[1]", StringComparison.OrdinalIgnoreCase)
                && coverXml.Contains("<w:date", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasExpectedBorder(BorderType? border)
        {
            if (border == null)
            {
                return false;
            }

            var value = border.GetAttribute("val", W.NamespaceName).Value ?? string.Empty;
            var size = border.Size?.Value;
            var color = border.Color?.Value ?? string.Empty;
            var themeColor = border.GetAttribute("themeColor", W.NamespaceName).Value ?? string.Empty;

            return string.Equals(value, "single", StringComparison.OrdinalIgnoreCase)
                && size == 24U
                && (string.Equals(color, "052F61", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(themeColor, "accent1", StringComparison.OrdinalIgnoreCase));
        }

        private static bool ContainsToday(string value, DateTime today)
        {
            var normalized = NormalizeText(value);
            var formats = new[]
            {
                "M/d/yyyy",
                "MM/dd/yyyy",
                "d/M/yyyy",
                "dd/MM/yyyy",
                "yyyy-MM-dd",
                "MMM d, yyyy",
                "MMMM d, yyyy",
                "d MMM yyyy",
                "d MMMM yyyy"
            };

            return formats.Any(format =>
                normalized.Contains(today.ToString(format, CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase)
                || normalized.Contains(today.ToString(format, CultureInfo.GetCultureInfo("en-US")), StringComparison.OrdinalIgnoreCase));
        }

        private static bool TryParseSupportedDate(string value, out DateTime parsedDate)
        {
            var formats = new[]
            {
                "M/d/yyyy",
                "MM/dd/yyyy",
                "M/d/yy",
                "MM/dd/yy",
                "d/M/yyyy",
                "dd/MM/yyyy",
                "d/M/yy",
                "dd/MM/yy",
                "yyyy-MM-dd",
                "dd-MM-yyyy",
                "d-M-yyyy",
                "MMM d, yyyy",
                "MMMM d, yyyy",
                "d MMM yyyy",
                "d MMMM yyyy"
            };

            foreach (var culture in new[] { CultureInfo.InvariantCulture, CultureInfo.GetCultureInfo("en-US"), CultureInfo.GetCultureInfo("en-GB") })
            {
                if (DateTime.TryParseExact(value, formats, culture, DateTimeStyles.None, out parsedDate))
                {
                    return true;
                }
            }

            return DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out parsedDate);
        }
    }
}
