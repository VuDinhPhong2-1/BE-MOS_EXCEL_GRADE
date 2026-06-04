using System.Xml.Linq;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project09
{
    public sealed class WP09T1Grader : IWordTaskGrader
    {
        private static readonly string[] CenteredStyleSetStyleIds =
        {
            "Normal",
            "Title",
            "Subtitle",
            "Heading1",
            "Heading2",
            "Heading3",
            "Quote",
            "IntenseQuote",
            "ListParagraph"
        };

        private static readonly Dictionary<string, string> StyleNameAliases = new(StringComparer.OrdinalIgnoreCase)
        {
            ["normal"] = "Normal",
            ["title"] = "Title",
            ["subtitle"] = "Subtitle",
            ["heading 1"] = "Heading1",
            ["heading1"] = "Heading1",
            ["heading 2"] = "Heading2",
            ["heading2"] = "Heading2",
            ["heading 3"] = "Heading3",
            ["heading3"] = "Heading3",
            ["quote"] = "Quote",
            ["intense quote"] = "IntenseQuote",
            ["intensequote"] = "IntenseQuote",
            ["list paragraph"] = "ListParagraph",
            ["listparagraph"] = "ListParagraph"
        };

        private const int MinimumCenteredLikeStyleCount = 6;

        public string TaskId => "W09-T01";
        public string TaskName => "Apply Centered style set";
        public decimal MaxScore => 1m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP09GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);

            if (!studentDocument.TryGetXmlPart("word/styles.xml", out var stylesXml))
            {
                WP09GraderHelpers.AddError(
                    result,
                    "Không tìm thấy word/styles.xml để xác minh bộ kiểu tài liệu.",
                    "Mở tài liệu trong Word, vào Design > Document Formatting, chọn style set Centered rồi lưu lại dưới định dạng .docx.");
                return result;
            }

            if (ContainsCenteredMetadata(stylesXml) || HasCenteredStyleSetCharacteristics(stylesXml, result))
            {
                result.Details.Add("Tài liệu có dấu hiệu đã áp dụng style set Centered trong word/styles.xml.");
                return result;
            }

            WP09GraderHelpers.AddError(
                result,
                "Không tìm thấy metadata hoặc đặc điểm style set Centered trong word/styles.xml.",
                "Vào Design > Document Formatting, chọn style set Centered cho toàn tài liệu, sau đó lưu file .docx và nộp lại.");
            return result;
        }

        private static bool ContainsCenteredMetadata(XDocument stylesXml)
        {
            var stylesText = WP09GraderHelpers.NormalizeText(stylesXml.ToString(SaveOptions.DisableFormatting));

            return stylesText.Contains("Centered", StringComparison.OrdinalIgnoreCase)
                || stylesText.Contains("StyleSet-Centered", StringComparison.OrdinalIgnoreCase)
                || stylesText.Contains("style set centered", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasCenteredStyleSetCharacteristics(XDocument stylesXml, TaskResult result)
        {
            var styleSignatures = BuildStyleSignatures(stylesXml);
            var centeredLikeStyles = styleSignatures.Count(pair => IsCenteredLikeStyle(pair.Key, pair.Value));
            var checkedStyles = CenteredStyleSetStyleIds.Count(styleSignatures.ContainsKey);

            if (checkedStyles == 0)
            {
                result.Details.Add("Không tìm thấy các style chính cần kiểm tra trong word/styles.xml.");
                return false;
            }

            var styleMatchRatio = (double)centeredLikeStyles / checkedStyles;
            result.Details.Add($"Kiểm tra đặc điểm Centered: {centeredLikeStyles}/{checkedStyles} style chính có căn giữa ({styleMatchRatio:P0}).");

            return centeredLikeStyles >= MinimumCenteredLikeStyleCount;
        }

        private static bool IsCenteredLikeStyle(string styleId, HashSet<string> signature)
        {
            return styleId.Equals("Title", StringComparison.OrdinalIgnoreCase)
                || styleId.Equals("Subtitle", StringComparison.OrdinalIgnoreCase)
                || signature.Any(token => token.Contains("jc[val=center]", StringComparison.OrdinalIgnoreCase));
        }

        private static Dictionary<string, HashSet<string>> BuildStyleSignatures(XDocument stylesXml)
        {
            var signatures = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (var style in stylesXml.Descendants(WP09GraderHelpers.W + "style"))
            {
                var canonicalStyleId = GetCanonicalStyleId(style);
                if (canonicalStyleId == null)
                {
                    continue;
                }

                signatures[canonicalStyleId] = BuildStyleSignature(style);
            }

            return signatures;
        }

        private static string? GetCanonicalStyleId(XElement style)
        {
            var styleId = style.Attribute(WP09GraderHelpers.W + "styleId")?.Value;
            if (!string.IsNullOrWhiteSpace(styleId)
                && CenteredStyleSetStyleIds.Contains(styleId, StringComparer.OrdinalIgnoreCase))
            {
                return CenteredStyleSetStyleIds.First(id => id.Equals(styleId, StringComparison.OrdinalIgnoreCase));
            }

            var styleName = style.Element(WP09GraderHelpers.W + "name")?.Attribute(WP09GraderHelpers.W + "val")?.Value;
            var normalizedStyleName = WP09GraderHelpers.NormalizeText(styleName).ToLowerInvariant();

            return StyleNameAliases.TryGetValue(normalizedStyleName, out var canonicalStyleName)
                ? canonicalStyleName
                : null;
        }

        private static HashSet<string> BuildStyleSignature(XElement style)
        {
            var signature = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            AddAttributeSignature(signature, "basedOn", style.Element(WP09GraderHelpers.W + "basedOn"));
            AddAttributeSignature(signature, "next", style.Element(WP09GraderHelpers.W + "next"));

            if (style.Element(WP09GraderHelpers.W + "qFormat") != null)
            {
                signature.Add("qFormat=true");
            }

            AddChildElementSignature(signature, "pPr", style.Element(WP09GraderHelpers.W + "pPr"), new[]
            {
                WP09GraderHelpers.W + "spacing",
                WP09GraderHelpers.W + "jc",
                WP09GraderHelpers.W + "outlineLvl"
            });

            AddChildElementSignature(signature, "rPr", style.Element(WP09GraderHelpers.W + "rPr"), new[]
            {
                WP09GraderHelpers.W + "rFonts",
                WP09GraderHelpers.W + "sz",
                WP09GraderHelpers.W + "szCs",
                WP09GraderHelpers.W + "color",
                WP09GraderHelpers.W + "b",
                WP09GraderHelpers.W + "bCs",
                WP09GraderHelpers.W + "i",
                WP09GraderHelpers.W + "iCs"
            });

            return signature;
        }

        private static void AddAttributeSignature(HashSet<string> signature, string tokenName, XElement? element)
        {
            var value = element?.Attribute(WP09GraderHelpers.W + "val")?.Value;
            if (!string.IsNullOrWhiteSpace(value))
            {
                signature.Add($"{tokenName}={value}");
            }
        }

        private static void AddChildElementSignature(
            HashSet<string> signature,
            string groupName,
            XElement? parentElement,
            IEnumerable<XName> childNames)
        {
            if (parentElement == null)
            {
                return;
            }

            foreach (var childName in childNames)
            {
                var childElement = parentElement.Element(childName);
                var normalizedElement = NormalizeStyleElement(childElement);
                if (!string.IsNullOrWhiteSpace(normalizedElement))
                {
                    signature.Add($"{groupName}.{childName.LocalName}={normalizedElement}");
                }
            }
        }

        private static string NormalizeStyleElement(XElement? element)
        {
            if (element == null)
            {
                return string.Empty;
            }

            var attributes = element
                .Attributes()
                .Where(attribute => !attribute.IsNamespaceDeclaration)
                .OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal)
                .ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal)
                .Select(attribute => $"{attribute.Name.LocalName}={WP09GraderHelpers.NormalizeText(attribute.Value)}");

            return $"{element.Name.LocalName}[{string.Join("|", attributes)}]";
        }
    }
}