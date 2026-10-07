using System.Xml.Linq;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public partial class XmlGradingRuleService
    {
        private static SpecialConditionEvalOutcome EvaluatePptPictureCropShape(PptPictureCropShapeConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptPictureCropShapeConfig.");

            if (!TryResolveTargetSlide(package, config.Slide, out var slide, out var err))
                return FailSpecialCondition(err ?? "Không tìm thấy slide mục tiêu.");

            var shapes = FindSlideShapes(slide.Doc, config.Shape ?? new PptShapeSelector { ShapeType = "picture" });
            if (shapes.Count == 0)
                return FailSpecialCondition($"Không tìm thấy hình ảnh mục tiêu trên Slide {slide.Index}.");

            var expectedPreset = string.IsNullOrWhiteSpace(config.ExpectedShapePreset) ? "ellipse" : config.ExpectedShapePreset.Trim().ToLowerInvariant();
            if (expectedPreset == "oval") expectedPreset = "ellipse";

            bool matched = false;
            foreach (var sp in shapes)
            {
                var prstGeom = sp.Descendants(ANs + "prstGeom").FirstOrDefault();
                var prst = prstGeom?.Attribute("prst")?.Value?.ToLowerInvariant();
                if (prst == expectedPreset)
                {
                    matched = true;
                    break;
                }
            }

            if (!matched)
            {
                return FailSpecialCondition(
                    $"Hình ảnh trên Slide {slide.Index} chưa được cắt xén (Crop) thành hình {config.ExpectedShapePreset}.",
                    fixAction: $"Chọn hình ảnh trên Slide {slide.Index} -> Picture Tools Format -> Crop -> Crop to Shape -> Chọn {config.ExpectedShapePreset}.");
            }

            return PassSpecialCondition($"Đã cắt xén hình ảnh thành công thành {config.ExpectedShapePreset} trên Slide {slide.Index}.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptShapeSize(PptShapeSizeConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptShapeSizeConfig.");

            if (!TryResolveTargetSlide(package, config.Slide, out var slide, out var err))
                return FailSpecialCondition(err ?? "Không tìm thấy slide mục tiêu.");

            var shapes = FindSlideShapes(slide.Doc, config.Shape);
            if (shapes.Count == 0)
                return FailSpecialCondition($"Không tìm thấy đối tượng trên Slide {slide.Index}.");

            var tolerance = config.ToleranceInches ?? 0.20;
            bool matched = false;
            string actualDetails = string.Empty;

            foreach (var sp in shapes)
            {
                var ext = sp.Descendants(ANs + "ext").FirstOrDefault();
                if (ext == null) continue;

                var cxStr = ext.Attribute("cx")?.Value;
                var cyStr = ext.Attribute("cy")?.Value;

                if (long.TryParse(cxStr, out var cxEmu) && long.TryParse(cyStr, out var cyEmu))
                {
                    var actualWidth = EmuToInches(cxEmu);
                    var actualHeight = EmuToInches(cyEmu);
                    actualDetails = $"Hiện tại: Height={actualHeight:F2}\", Width={actualWidth:F2}\"";

                    bool hOk = !config.ExpectedHeightInches.HasValue || InchesEqual(actualHeight, config.ExpectedHeightInches.Value, tolerance);
                    bool wOk = !config.ExpectedWidthInches.HasValue || InchesEqual(actualWidth, config.ExpectedWidthInches.Value, tolerance);

                    if (hOk && wOk)
                    {
                        matched = true;
                        break;
                    }
                }
            }

            if (!matched)
            {
                return FailSpecialCondition(
                    $"Kích thước đối tượng trên Slide {slide.Index} chưa chính xác. {actualDetails}.",
                    fixAction: $"Thay đổi kích thước Height={config.ExpectedHeightInches}\", Width={config.ExpectedWidthInches}\" trong Picture/Drawing Tools Format -> Size.");
            }

            return PassSpecialCondition($"Đã định dạng kích thước chính xác trên Slide {slide.Index}.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptShapeGroup(PptShapeGroupConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptShapeGroupConfig.");

            if (!TryResolveTargetSlide(package, config.Slide, out var slide, out var err))
                return FailSpecialCondition(err ?? "Không tìm thấy slide mục tiêu.");

            var grpSp = slide.Doc.Descendants(PNs + "grpSp").FirstOrDefault();
            if (grpSp == null)
            {
                return FailSpecialCondition(
                    $"Các hình trên Slide {slide.Index} chưa được nhóm (Group) lại với nhau.",
                    fixAction: "Chọn tất cả các hình -> Drawing Tools Format -> Arrange -> Group -> Group.");
            }

            var childShapes = grpSp.Elements().Where(e => e.Name == PNs + "sp" || e.Name == PNs + "pic").ToList();
            if (config.MinChildCount.HasValue && childShapes.Count < config.MinChildCount.Value)
            {
                return FailSpecialCondition($"Số lượng hình trong nhóm chưa đủ (hiện có {childShapes.Count} hình).");
            }

            return PassSpecialCondition($"Đã nhóm các hình lại thành công trên Slide {slide.Index}.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptPictureStyle(PptPictureStyleConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptPictureStyleConfig.");

            if (!TryResolveTargetSlide(package, config.Slide, out var slide, out var err))
                return FailSpecialCondition(err ?? "Không tìm thấy slide mục tiêu.");

            var pictures = FindSlideShapes(slide.Doc, config.Shape ?? new PptShapeSelector { ShapeType = "picture" });
            if (pictures.Count == 0)
                return FailSpecialCondition($"Không tìm thấy hình ảnh trên Slide {slide.Index}.");

            bool hasStyle = false;
            foreach (var pic in pictures)
            {
                // Kiểm tra hiệu ứng 3D bevel hoặc viền hoặc style
                var sp3d = pic.Descendants(ANs + "sp3d").FirstOrDefault();
                var scene3d = pic.Descendants(ANs + "scene3d").FirstOrDefault();
                var ln = pic.Descendants(ANs + "ln").FirstOrDefault();
                var style = pic.Descendants(PNs + "style").FirstOrDefault();

                if (sp3d != null || scene3d != null || ln != null || style != null)
                {
                    hasStyle = true;
                    break;
                }
            }

            if (!hasStyle)
            {
                return FailSpecialCondition(
                    $"Chưa áp dụng kiểu {config.ExpectedStyleName} cho hình ảnh trên Slide {slide.Index}.",
                    fixAction: $"Chọn hình ảnh -> Picture Tools Format -> Picture Styles -> Chọn {config.ExpectedStyleName}.");
            }

            return PassSpecialCondition($"Đã áp dụng kiểu {config.ExpectedStyleName} thành công cho hình ảnh.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptShapeArrange(PptShapeArrangeConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptShapeArrangeConfig.");

            if (!TryResolveTargetSlide(package, config.Slide, out var slide, out var err))
                return FailSpecialCondition(err ?? "Không tìm thấy slide mục tiêu.");

            var spTree = slide.Doc.Descendants(PNs + "spTree").FirstOrDefault();
            if (spTree == null) return FailSpecialCondition("Không tìm thấy spTree trên slide.");

            var shapes = FindSlideShapes(slide.Doc, config.Shape);
            if (shapes.Count == 0)
                return FailSpecialCondition($"Không tìm thấy đối tượng cần căn chỉnh trên Slide {slide.Index}.");

            // Xác định đối tượng mục tiêu:
            // Nếu không chỉ định shape cụ thể và có nhiều hình ảnh (pic),
            // ưu tiên chọn hình ảnh ở giữa (dựa theo hoành độ X)
            XElement targetShape;
            var pics = shapes.Where(e => e.Name == PNs + "pic").ToList();
            if (config.Shape == null && pics.Count > 1)
            {
                targetShape = pics
                    .OrderBy(p => GetShapeCoord(p, "x"))
                    .Skip(pics.Count / 2)
                    .FirstOrDefault() ?? shapes.First();
            }
            else
            {
                targetShape = shapes.First();
            }

            // Lấy chiều cao của slide để kiểm tra căn giữa dọc
            long slideHeight = 6858000; // Mặc định chuẩn 16:9 (6,858,000 EMU)
            if (package.XmlParts.TryGetValue("ppt/presentation.xml", out var presXml))
            {
                try
                {
                    var presDoc = XDocument.Parse(presXml);
                    var sldSz = presDoc.Descendants(PNs + "sldSz").FirstOrDefault();
                    if (sldSz != null && long.TryParse(sldSz.Attribute("cy")?.Value, out var cy) && cy > 0)
                    {
                        slideHeight = cy;
                    }
                }
                catch { }
            }

            if (config.CheckAlignMiddle == true)
            {
                var y = GetShapeCoord(targetShape, "y");
                var cy = GetShapeDimension(targetShape, "cy");
                if (cy > 0)
                {
                    var expectedY = (slideHeight - cy) / 2;
                    var centerY = y + cy / 2;
                    var slideCenterY = slideHeight / 2;

                    // Dung sai 150,000 EMU (~ 0.16 inches)
                    bool isMiddleAligned = Math.Abs(y - expectedY) <= 150_000 || Math.Abs(centerY - slideCenterY) <= 150_000;
                    if (!isMiddleAligned)
                    {
                        return FailSpecialCondition(
                            $"Hình ảnh trên Slide {slide.Index} chưa được căn giữa theo chiều dọc (Align Middle).",
                            fixAction: "Chọn hình ảnh ở giữa -> Picture Format -> Align -> Align Middle.");
                    }
                }
            }

            if (config.CheckBringToFront == true)
            {
                var elements = spTree.Elements()
                    .Where(e => e.Name == PNs + "sp" || e.Name == PNs + "pic" || e.Name == PNs + "grpSp" || e.Name == PNs + "graphicFrame")
                    .ToList();

                if (elements.Count > 0)
                {
                    var targetIdx = elements.IndexOf(targetShape);
                    var picElements = elements.Where(e => e.Name == PNs + "pic").ToList();
                    bool isTopPic = picElements.Count > 0 && picElements.Last() == targetShape;
                    bool isNearEnd = targetIdx >= elements.Count - 2;

                    if (!isTopPic && !isNearEnd)
                    {
                        return FailSpecialCondition(
                            $"Hình ảnh trên Slide {slide.Index} chưa được đưa lên phía trước cùng (Bring to Front).",
                            fixAction: "Chọn hình ảnh -> Picture Format -> Arrange -> Bring Forward -> Bring to Front.");
                    }
                }
            }

            return PassSpecialCondition($"Đã căn chỉnh vị trí đối tượng thành công trên Slide {slide.Index}.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptTextBox(PptTextBoxConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptTextBoxConfig.");

            if (!TryResolveTargetSlide(package, config.Slide, out var slide, out var err))
                return FailSpecialCondition(err ?? "Không tìm thấy slide mục tiêu.");

            var textBoxes = slide.Doc.Descendants(PNs + "sp")
                .Where(sp =>
                {
                    var cNvSpPr = sp.Descendants(PNs + "cNvSpPr").FirstOrDefault();
                    return cNvSpPr?.Attribute("txBox")?.Value == "1" || sp.Descendants(ANs + "txBody").Any();
                }).ToList();

            var targetNorm = NormalizePlainText(config.ExpectedText);
            XElement? matchedBox = null;

            foreach (var tb in textBoxes)
            {
                var text = NormalizePlainText(GetElementText(tb));
                if (text.Contains(targetNorm, StringComparison.OrdinalIgnoreCase))
                {
                    matchedBox = tb;
                    break;
                }
            }

            if (matchedBox == null)
            {
                return FailSpecialCondition(
                    $"Không tìm thấy hộp văn bản (Text Box) chứa nội dung '{config.ExpectedText}' trên Slide {slide.Index}.",
                    fixAction: $"Chèn Text Box mới vào góc dưới bên phải Slide {slide.Index} và nhập nội dung '{config.ExpectedText}'.");
            }

            if (config.ExpectedWidthInches.HasValue)
            {
                var ext = matchedBox.Descendants(ANs + "ext").FirstOrDefault();
                var cxStr = ext?.Attribute("cx")?.Value;
                if (long.TryParse(cxStr, out var cxEmu))
                {
                    var width = EmuToInches(cxEmu);
                    if (!InchesEqual(width, config.ExpectedWidthInches.Value, 0.35))
                    {
                        return FailSpecialCondition(
                            $"Chiều rộng của Text Box chưa đúng {config.ExpectedWidthInches.Value}\" (hiện tại: {width:F2}\").",
                            fixAction: $"Chọn Text Box -> Drawing Tools Format -> Size -> Shape Width: {config.ExpectedWidthInches.Value}\".");
                    }
                }
            }

            return PassSpecialCondition($"Đã chèn hộp văn bản chính xác trên Slide {slide.Index}.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptSlideLayout(PptSlideLayoutConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptSlideLayoutConfig.");

            if (!TryResolveTargetSlide(package, config.Slide, out var slide, out var err))
                return FailSpecialCondition(err ?? "Không tìm thấy slide mục tiêu.");

            var relsPath = GetSlideRelsPath(slide.Path);
            if (!package.XmlParts.TryGetValue(relsPath, out var relsXml))
            {
                return FailSpecialCondition($"Không tìm thấy file quan hệ cho Slide {slide.Index}.");
            }

            var layoutPart = ResolveLayoutPartPath(relsXml);
            if (string.IsNullOrWhiteSpace(layoutPart) || !package.XmlParts.TryGetValue(layoutPart, out var layoutXml))
            {
                return FailSpecialCondition($"Không tìm thấy layout của Slide {slide.Index}.");
            }

            var expectedNorm = NormalizePlainText(config.ExpectedLayoutName);
            var layoutDoc = XDocument.Parse(layoutXml);
            var cSld = layoutDoc.Descendants(PNs + "cSld").FirstOrDefault();
            var layoutName = cSld?.Attribute("name")?.Value ?? string.Empty;

            if (!NormalizePlainText(layoutName).Contains(expectedNorm, StringComparison.OrdinalIgnoreCase))
            {
                return FailSpecialCondition(
                    $"Bố cục của Slide {slide.Index} chưa đúng '{config.ExpectedLayoutName}' (hiện tại: '{layoutName}').",
                    fixAction: $"Chọn Slide {slide.Index} -> Home -> Layout -> Chọn {config.ExpectedLayoutName}.");
            }

            return PassSpecialCondition($"Bố cục của Slide {slide.Index} đã được đổi thành '{config.ExpectedLayoutName}'.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptShapeStyle(PptShapeStyleConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptShapeStyleConfig.");

            if (!TryResolveTargetSlide(package, config.Slide, out var slide, out var err))
                return FailSpecialCondition(err ?? "Không tìm thấy slide mục tiêu.");

            var shapes = FindSlideShapes(slide.Doc, config.Shape);
            if (shapes.Count == 0)
                return FailSpecialCondition($"Không tìm thấy hình dạng trên Slide {slide.Index}.");

            bool hasStyle = false;
            foreach (var sp in shapes)
            {
                var style = sp.Descendants(PNs + "style").FirstOrDefault();
                if (style != null)
                {
                    hasStyle = true;
                    break;
                }
            }

            if (!hasStyle)
            {
                return FailSpecialCondition(
                    $"Hình dạng trên Slide {slide.Index} chưa được áp dụng kiểu {config.ExpectedStyleName}.",
                    fixAction: $"Chọn hình -> Drawing Tools Format -> Shape Styles -> Chọn {config.ExpectedStyleName}.");
            }

            return PassSpecialCondition($"Đã áp dụng kiểu {config.ExpectedStyleName} thành công trên Slide {slide.Index}.");
        }

        private static string GetSlideRelsPath(string slidePath)
        {
            var fileName = Path.GetFileName(slidePath);
            return $"ppt/slides/_rels/{fileName}.rels";
        }

        private static string? ResolveLayoutPartPath(string relsXml)
        {
            try
            {
                var doc = XDocument.Parse(relsXml);
                foreach (var rel in doc.Descendants(RelsNs + "Relationship"))
                {
                    var type = rel.Attribute("Type")?.Value;
                    if (type != null && type.EndsWith("slideLayout", StringComparison.OrdinalIgnoreCase))
                    {
                        var target = rel.Attribute("Target")?.Value;
                        if (!string.IsNullOrWhiteSpace(target))
                        {
                            var clean = target.Replace("../", "").TrimStart('/');
                            return "ppt/" + clean;
                        }
                    }
                }
            }
            catch
            {
            }
            return null;
        }
    }
}
