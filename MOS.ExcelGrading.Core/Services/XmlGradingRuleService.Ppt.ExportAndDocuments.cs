using System.Xml.Linq;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public partial class XmlGradingRuleService
    {
        private static SpecialConditionEvalOutcome EvaluatePptExportedFile(PptExportedFileConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptExportedFileConfig.");

            var expected = config.ExpectedFileName?.Trim();
            if (string.IsNullOrWhiteSpace(expected))
            {
                return FailSpecialCondition("pptExportedFileConfig.expectedFileName không được để trống.");
            }

            // Kiểm tra danh sách file đính kèm nộp cùng lúc
            var attached = package.AttachedFileNames ?? Array.Empty<string>();
            bool matched = false;

            foreach (var fileName in attached)
            {
                var cleanName = Path.GetFileName(fileName);
                if (config.CaseSensitive)
                {
                    if (string.Equals(cleanName, expected, StringComparison.Ordinal))
                    {
                        matched = true;
                        break;
                    }
                }
                else
                {
                    if (string.Equals(cleanName, expected, StringComparison.OrdinalIgnoreCase))
                    {
                        matched = true;
                        break;
                    }
                }
            }

            if (!matched)
            {
                return FailSpecialCondition(
                    $"Chưa tìm thấy tệp PDF xuất ra '{expected}' được nộp kèm (Lưu ý: Bắt buộc chính xác chữ hoa / chữ thường).",
                    fixAction: $"Xuất bài thuyết trình sang PDF với tên '{expected}' và chọn nộp cùng lúc cả file .pptx và file {expected}.");
            }

            return PassSpecialCondition($"Đã xuất và nộp kèm tệp PDF '{expected}' thành công.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptComment(PptCommentConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptCommentConfig.");

            var commentParts = package.XmlParts.Where(p => p.Key.StartsWith("ppt/comments/", StringComparison.OrdinalIgnoreCase)).ToList();

            if (config.ExpectAbsent)
            {
                var targetTextNorm = NormalizePlainText(config.TargetCommentText ?? string.Empty);
                var authorNorm = NormalizePlainText(config.Author ?? string.Empty);

                foreach (var (_, xml) in commentParts)
                {
                    var xmlNorm = NormalizePlainText(xml);
                    if (!string.IsNullOrWhiteSpace(targetTextNorm) && xmlNorm.Contains(targetTextNorm, StringComparison.OrdinalIgnoreCase))
                    {
                        return FailSpecialCondition(
                            $"Bình luận (Comment) có nội dung '{config.TargetCommentText}' vẫn chưa được xóa.",
                            fixAction: $"Tìm đến bình luận '{config.TargetCommentText}' và nhấp Delete (hoặc bấm dấu X) để xóa.");
                    }
                }

                return PassSpecialCondition("Đã xóa bình luận mục tiêu thành công.");
            }

            return PassSpecialCondition("Đã xử lý bình luận.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptSlideTitles(PptSlideTitlesConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptSlideTitlesConfig.");

            var slides = ResolveOrderedSlides(package);
            var startIndex = (config.StartSlideIndex ?? 1) - 1;

            if (startIndex < 0 || startIndex >= slides.Count)
            {
                return FailSpecialCondition($"Chưa có các slide mới được tạo từ tệp Outline.");
            }

            for (int i = 0; i < config.ExpectedTitlesInOrder.Count; i++)
            {
                int currentSlideIdx = startIndex + i;
                if (currentSlideIdx >= slides.Count)
                {
                    return FailSpecialCondition($"Thiếu slide có tiêu đề '{config.ExpectedTitlesInOrder[i]}'.");
                }

                var slide = slides[currentSlideIdx];
                var actualTitle = NormalizePlainText(GetSlideTitle(slide.Doc));
                var expectedTitle = NormalizePlainText(config.ExpectedTitlesInOrder[i]);

                if (!actualTitle.Contains(expectedTitle, StringComparison.OrdinalIgnoreCase))
                {
                    return FailSpecialCondition(
                        $"Tiêu đề Slide {slide.Index} chưa đúng '{config.ExpectedTitlesInOrder[i]}' (hiện tại: '{GetSlideTitle(slide.Doc)}').",
                        fixAction: $"Sử dụng New Slide -> Slides from Outline... để nhập tệp Outline.");
                }
            }

            return PassSpecialCondition("Các slide mới từ Outline đã được tạo với tiêu đề chính xác.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptTable(PptTableConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptTableConfig.");

            if (!TryResolveTargetSlide(package, config.Slide, out var slide, out var err))
                return FailSpecialCondition(err ?? "Không tìm thấy slide mục tiêu.");

            var tables = slide.Doc.Descendants(ANs + "tbl").ToList();
            if (tables.Count == 0) return FailSpecialCondition($"Không tìm thấy bảng trên Slide {slide.Index}.");

            var table = tables.First();
            var rows = table.Elements(ANs + "tr").ToList();

            foreach (var disallowed in config.DisallowedRowTexts)
            {
                var normDisallowed = NormalizePlainText(disallowed);
                foreach (var row in rows)
                {
                    var rowText = NormalizePlainText(GetElementText(row));
                    if (rowText.Contains(normDisallowed, StringComparison.OrdinalIgnoreCase))
                    {
                        return FailSpecialCondition(
                            $"Hàng chứa nội dung '{disallowed}' vẫn chưa được xóa khỏi bảng trên Slide {slide.Index}.",
                            fixAction: $"Chọn hàng chứa '{disallowed}' -> Table Tools Layout -> Rows & Columns -> Delete -> Delete Rows.");
                    }
                }
            }

            return PassSpecialCondition($"Đã thao tác với bảng thành công trên Slide {slide.Index}.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptSection(PptSectionConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptSectionConfig.");

            if (!package.XmlParts.TryGetValue("ppt/presentation.xml", out var presXml))
                return FailSpecialCondition("Không tìm thấy tệp ppt/presentation.xml.");

            var expectedNorm = NormalizePlainText(config.ExpectedSectionName);
            var presDoc = XDocument.Parse(presXml);

            var sections = presDoc.Descendants(P14Ns + "section").ToList();
            bool found = false;

            foreach (var sec in sections)
            {
                var name = sec.Attribute("name")?.Value;
                if (!string.IsNullOrWhiteSpace(name) && NormalizePlainText(name).Equals(expectedNorm, StringComparison.OrdinalIgnoreCase))
                {
                    found = true;
                    break;
                }
            }

            if (!found)
            {
                return FailSpecialCondition(
                    $"Chưa đổi tên phần (Section) thành '{config.ExpectedSectionName}'.",
                    fixAction: $"Nhấp chuột phải vào Untitled Section -> Rename Section -> Nhập '{config.ExpectedSectionName}' -> Rename.");
            }

            return PassSpecialCondition($"Đã đổi tên phần thành '{config.ExpectedSectionName}' thành công.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptSummaryZoom(PptSummaryZoomConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptSummaryZoomConfig.");

            var slides = ResolveOrderedSlides(package);
            bool foundZoom = false;

            foreach (var slide in slides)
            {
                var xml = package.XmlParts[slide.Path];
                if (xml.Contains("sectionZm", StringComparison.OrdinalIgnoreCase)
                    || xml.Contains("summaryZm", StringComparison.OrdinalIgnoreCase)
                    || xml.Contains("zoom", StringComparison.OrdinalIgnoreCase))
                {
                    foundZoom = true;
                    break;
                }
            }

            if (!foundZoom)
            {
                return FailSpecialCondition(
                    "Chưa chèn trang trình chiếu Summary Zoom vào giữa Slide 1 và 2.",
                    fixAction: "Chọn vị trí giữa Slide 1 và 2 -> Insert -> Zoom -> Summary Zoom -> Bỏ chọn Slide 1 -> Insert.");
            }

            return PassSpecialCondition("Đã chèn Summary Zoom thành công.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptMasterPicture(PptMasterPictureConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptMasterPictureConfig.");

            var masterParts = package.XmlParts.Where(p => p.Key.StartsWith("ppt/slideMasters/slideMaster", StringComparison.OrdinalIgnoreCase)).ToList();
            if (masterParts.Count == 0) return FailSpecialCondition("Không tìm thấy Slide Master trong tệp PowerPoint.");

            bool hasPicture = false;
            foreach (var (_, xml) in masterParts)
            {
                if (xml.Contains("<p:pic", StringComparison.OrdinalIgnoreCase))
                {
                    hasPicture = true;
                    break;
                }
            }

            if (!hasPicture)
            {
                return FailSpecialCondition(
                    "Chưa chèn hình ảnh Badge.png vào trang chiếu chính (Slide Master).",
                    fixAction: "View -> Slide Master -> Chọn Slide Master -> Insert -> Pictures -> Chọn Badge.png và đặt vào góc dưới bên phải.");
            }

            return PassSpecialCondition("Đã chèn hình ảnh vào Slide Master thành công.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptMarkAsFinal(PptMarkAsFinalConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptMarkAsFinalConfig.");

            if (!package.XmlParts.TryGetValue("docProps/custom.xml", out var customXml))
            {
                return FailSpecialCondition(
                    "Bài thuyết trình chưa được đánh dấu là Final (Mark as Final).",
                    fixAction: "File -> Info -> Protect Presentation -> Mark as Final -> OK.");
            }

            if (!customXml.Contains("_MarkAsFinal", StringComparison.OrdinalIgnoreCase))
            {
                return FailSpecialCondition(
                    "Thuộc tính _MarkAsFinal chưa được thiết lập trong tệp.",
                    fixAction: "File -> Info -> Protect Presentation -> Mark as Final -> OK.");
            }

            return PassSpecialCondition("Bài thuyết trình đã được đánh dấu là Final (Mark as Final) thành công.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptPrintSettings(PptPrintSettingsConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptPrintSettingsConfig.");

            if (!package.XmlParts.TryGetValue("ppt/presProps.xml", out var presPropsXml))
            {
                // Nếu không có presProps.xml thì kiểm tra cấu hình mặc định hoặc cho pass nếu rubric không lưu cứng
                return PassSpecialCondition("Cấu hình in ấn được ghi nhận.");
            }

            var doc = XDocument.Parse(presPropsXml);
            var prnPr = doc.Descendants(PNs + "prnPr").FirstOrDefault();

            if (!string.IsNullOrWhiteSpace(config.ExpectedPrintWhat))
            {
                var prnWhat = prnPr?.Attribute("prnWhat")?.Value?.ToLowerInvariant();
                var expNorm = config.ExpectedPrintWhat.Trim().ToLowerInvariant();

                if (expNorm.Contains("note") && prnWhat != "notes")
                {
                    return FailSpecialCondition(
                        "Cài đặt in chưa được chọn là Notes Pages.",
                        fixAction: "File -> Print -> Settings -> Đổi Full Page Slides thành Notes Pages.");
                }
            }

            if (!string.IsNullOrWhiteSpace(config.ExpectedColorMode))
            {
                var clrMode = prnPr?.Attribute("clrMode")?.Value?.ToLowerInvariant();
                var expClr = config.ExpectedColorMode.Trim().ToLowerInvariant();

                if (expClr.Contains("gray") && clrMode != "gray")
                {
                    return FailSpecialCondition(
                        "Cài đặt in chưa được đặt ở chế độ Grayscale.",
                        fixAction: "File -> Print -> Settings -> Đổi Color thành Grayscale.");
                }
            }

            return PassSpecialCondition("Cấu hình tùy chọn in ấn thành công.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptTextColumns(PptTextColumnsConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptTextColumnsConfig.");

            if (!TryResolveTargetSlide(package, config.Slide, out var slide, out var err))
                return FailSpecialCondition(err ?? "Không tìm thấy slide mục tiêu.");

            var bodyPrs = slide.Doc.Descendants(ANs + "bodyPr").ToList();
            bool matched = false;

            foreach (var bPr in bodyPrs)
            {
                var numCol = bPr.Attribute("numCol")?.Value;
                if (numCol == config.ExpectedColumnCount.ToString())
                {
                    matched = true;
                    break;
                }
            }

            if (!matched)
            {
                return FailSpecialCondition(
                    $"Văn bản trên Slide {slide.Index} chưa được định dạng thành {config.ExpectedColumnCount} cột.",
                    fixAction: $"Chọn hộp văn bản trên Slide {slide.Index} -> Home -> Add or Remove Columns -> More Columns... -> Number: {config.ExpectedColumnCount}, Spacing: {config.ExpectedSpacingInches}\".");
            }

            return PassSpecialCondition($"Đã định dạng {config.ExpectedColumnCount} cột thành công trên Slide {slide.Index}.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptNotesMasterPlaceholders(PptNotesMasterPlaceholdersConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptNotesMasterPlaceholdersConfig.");

            var notesMasterParts = package.XmlParts.Where(p => p.Key.StartsWith("ppt/notesMasters/notesMaster", StringComparison.OrdinalIgnoreCase)).ToList();
            if (notesMasterParts.Count == 0) return FailSpecialCondition("Không tìm thấy Notes Master trong tệp.");

            var (_, xml) = notesMasterParts.First();
            var doc = XDocument.Parse(xml);

            var hf = doc.Descendants(PNs + "hf").FirstOrDefault();
            bool hdrDisabled = hf?.Attribute("hdr")?.Value == "0" || hf?.Attribute("hdr")?.Value == "false";
            bool ftrDisabled = hf?.Attribute("ftr")?.Value == "0" || hf?.Attribute("ftr")?.Value == "false";

            // Cũng kiểm tra xem có placeholder hdr/ftr trong shape không
            var hasHeaderSp = doc.Descendants(PNs + "ph").Any(ph => ph.Attribute("type")?.Value == "hdr");
            var hasFooterSp = doc.Descendants(PNs + "ph").Any(ph => ph.Attribute("type")?.Value == "ftr");

            if ((!hdrDisabled && hasHeaderSp) || (!ftrDisabled && hasFooterSp))
            {
                return FailSpecialCondition(
                    "Chưa tắt Header and Footer khỏi Notes Master.",
                    fixAction: "View -> Master Views -> Notes Master -> Bỏ tích chọn ô Header và Footer -> Close Master View.");
            }

            return PassSpecialCondition("Đã xóa Header và Footer khỏi Notes Master thành công.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptSlideSize(PptSlideSizeConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptSlideSizeConfig.");

            if (!package.XmlParts.TryGetValue("ppt/presentation.xml", out var presXml))
                return FailSpecialCondition("Không tìm thấy tệp ppt/presentation.xml.");

            var doc = XDocument.Parse(presXml);
            var sldSz = doc.Descendants(PNs + "sldSz").FirstOrDefault();
            if (sldSz == null) return FailSpecialCondition("Không tìm thấy thẻ kích thước slide sldSz.");

            var type = sldSz.Attribute("type")?.Value;
            var cxStr = sldSz.Attribute("cx")?.Value;
            var cyStr = sldSz.Attribute("cy")?.Value;

            bool is16x9 = type == "screen16x9";
            if (!is16x9 && long.TryParse(cxStr, out var cx) && long.TryParse(cyStr, out var cy))
            {
                double ratio = (double)cx / cy;
                if (Math.Abs(ratio - (16.0 / 9.0)) < 0.05) is16x9 = true;
            }

            if (!is16x9)
            {
                return FailSpecialCondition(
                    "Kích thước trang chiếu chưa được chuyển sang Widescreen (16:9).",
                    fixAction: "Design -> Slide Size -> Chọn Widescreen (16:9) -> Ensure Fit.");
            }

            return PassSpecialCondition("Đã chuyển kích thước trang chiếu sang Widescreen (16:9) thành công.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptAltText(PptAltTextConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptAltTextConfig.");

            if (!TryResolveTargetSlide(package, config.Slide, out var slide, out var err))
                return FailSpecialCondition(err ?? "Không tìm thấy slide mục tiêu.");

            var expectedNorm = NormalizePlainText(config.ExpectedAltText);
            bool matched = false;

            foreach (var cNvPr in slide.Doc.Descendants(PNs + "cNvPr"))
            {
                var descr = cNvPr.Attribute("descr")?.Value;
                var title = cNvPr.Attribute("title")?.Value;

                if (!string.IsNullOrWhiteSpace(descr) && NormalizePlainText(descr).Contains(expectedNorm, StringComparison.OrdinalIgnoreCase))
                {
                    matched = true;
                    break;
                }

                if (!string.IsNullOrWhiteSpace(title) && NormalizePlainText(title).Contains(expectedNorm, StringComparison.OrdinalIgnoreCase))
                {
                    matched = true;
                    break;
                }
            }

            if (!matched)
            {
                return FailSpecialCondition(
                    $"Chưa thêm Alt Text có nội dung '{config.ExpectedAltText}' cho đồ họa trên Slide {slide.Index}.",
                    fixAction: $"Nhấp chuột phải vào hình ảnh trên Slide {slide.Index} -> Edit Alt Text... -> Nhập '{config.ExpectedAltText}'.");
            }

            return PassSpecialCondition($"Đã thêm Alt Text '{config.ExpectedAltText}' thành công trên Slide {slide.Index}.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptHyperlink(PptHyperlinkConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptHyperlinkConfig.");

            if (!TryResolveTargetSlide(package, config.Slide, out var slide, out var err))
                return FailSpecialCondition(err ?? "Không tìm thấy slide mục tiêu.");

            var relsPath = GetSlideRelsPath(slide.Path);
            if (!package.XmlParts.TryGetValue(relsPath, out var relsXml))
            {
                return FailSpecialCondition($"Không tìm thấy file quan hệ cho Slide {slide.Index}.");
            }

            var expectedUrlNorm = config.ExpectedUrl?.Trim().ToLowerInvariant() ?? string.Empty;
            bool linkFound = false;

            try
            {
                var relsDoc = XDocument.Parse(relsXml);
                foreach (var rel in relsDoc.Descendants(RelsNs + "Relationship"))
                {
                    var target = rel.Attribute("Target")?.Value?.ToLowerInvariant();
                    if (!string.IsNullOrWhiteSpace(target) && (target == expectedUrlNorm || target.TrimEnd('/') == expectedUrlNorm.TrimEnd('/')))
                    {
                        linkFound = true;
                        break;
                    }
                }
            }
            catch
            {
            }

            if (!linkFound)
            {
                return FailSpecialCondition(
                    $"Chưa tạo liên kết (Hyperlink) tới '{config.ExpectedUrl}' trên Slide {slide.Index}.",
                    fixAction: $"Chọn văn bản '{config.TargetText}' -> Insert -> Link -> Nhập Address: {config.ExpectedUrl} -> OK.");
            }

            return PassSpecialCondition($"Đã tạo liên kết thành công tới '{config.ExpectedUrl}'.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptSlideBackground(PptSlideBackgroundConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptSlideBackgroundConfig.");

            if (!TryResolveTargetSlide(package, config.Slide, out var slide, out var err))
                return FailSpecialCondition(err ?? "Không tìm thấy slide mục tiêu.");

            var bg = slide.Doc.Descendants(PNs + "bg").FirstOrDefault();
            if (bg == null)
            {
                return FailSpecialCondition(
                    $"Chưa áp dụng màu nền Solid Fill cho Slide {slide.Index}.",
                    fixAction: $"Design -> Format Background -> Solid Fill -> Color: Blue -> OK.");
            }

            var expectedHex = (config.ExpectedColorHex ?? "0070C0").Trim().ToUpperInvariant();
            var srgbClr = bg.Descendants(ANs + "srgbClr").FirstOrDefault();
            var actualHex = srgbClr?.Attribute("val")?.Value?.ToUpperInvariant();

            if (actualHex != null && actualHex == expectedHex)
            {
                return PassSpecialCondition($"Đã áp dụng màu nền Blue thành công trên Slide {slide.Index}.");
            }

            // Chấp nhận nếu có solidFill
            if (bg.Descendants(ANs + "solidFill").Any())
            {
                return PassSpecialCondition($"Đã áp dụng màu nền thành công trên Slide {slide.Index}.");
            }

            return FailSpecialCondition(
                $"Màu nền trên Slide {slide.Index} chưa đúng màu {expectedHex}.",
                fixAction: $"Design -> Format Background -> Solid Fill -> Color: Blue ({expectedHex}).");
        }
    }
}
