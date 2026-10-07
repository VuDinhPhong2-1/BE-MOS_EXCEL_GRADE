using System.Xml.Linq;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public partial class XmlGradingRuleService
    {
        private static SpecialConditionEvalOutcome EvaluatePptVideo(PptVideoConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptVideoConfig.");

            if (!TryResolveTargetSlide(package, config.Slide, out var slide, out var err))
                return FailSpecialCondition(err ?? "Không tìm thấy slide mục tiêu.");

            // Kiểm tra sự hiện diện của video trên slide (P02-T2)
            bool hasVideo = false;
            XElement? matchedMediaElement = null;

            // 1. Tìm trong pic hoặc graphicFrame có videoFile / media
            foreach (var pic in slide.Doc.Descendants(PNs + "pic"))
            {
                var videoFile = pic.Descendants(ANs + "videoFile").FirstOrDefault();
                var media14 = pic.Descendants(P14Ns + "media").FirstOrDefault();
                if (videoFile != null || media14 != null)
                {
                    hasVideo = true;
                    matchedMediaElement = media14 ?? videoFile;
                    break;
                }
            }

            // 2. Tìm qua file rels của slide xem có quan hệ kiểu video / media không
            if (!hasVideo)
            {
                var relsPath = GetSlideRelsPath(slide.Path);
                if (package.XmlParts.TryGetValue(relsPath, out var relsXml))
                {
                    if (relsXml.Contains("video", StringComparison.OrdinalIgnoreCase) || relsXml.Contains("media", StringComparison.OrdinalIgnoreCase))
                    {
                        hasVideo = true;
                    }
                }
            }

            if (!hasVideo)
            {
                return FailSpecialCondition(
                    $"Trên Slide {slide.Index} chưa có đối tượng Video / Screen Recording.",
                    fixAction: $"Chèn Screen Recording hoặc Video vào Slide {slide.Index}.");
            }

            // Kiểm tra Trim video (P02-T4)
            if (config.ExpectedTrimEndMs.HasValue || config.ExpectedTrimEndSeconds.HasValue)
            {
                long targetMs = config.ExpectedTrimEndMs ?? (long)(config.ExpectedTrimEndSeconds!.Value * 1000);
                var trim = slide.Doc.Descendants(P14Ns + "trim").FirstOrDefault();
                var endStr = trim?.Attribute("end")?.Value;

                if (string.IsNullOrWhiteSpace(endStr) || !long.TryParse(endStr, out var actualEndMs) || Math.Abs(actualEndMs - targetMs) > 1000)
                {
                    return FailSpecialCondition(
                        $"End Time của video trên Slide {slide.Index} chưa được cắt đúng ở giây thứ {targetMs / 1000} (14s).",
                        fixAction: "Chọn Video -> Video Tools Playback -> Trim Video -> Đặt End Time thành 00:14.");
                }
            }

            return PassSpecialCondition($"Đã xử lý video thành công trên Slide {slide.Index}.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptSlideTransition(PptSlideTransitionConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptSlideTransitionConfig.");

            var slides = ResolveOrderedSlides(package);
            if (slides.Count == 0) return FailSpecialCondition("Không tìm thấy slide nào trong bài trình chiếu.");

            var targetSlides = config.ApplyToAll == true
                ? slides
                : (TryResolveTargetSlide(package, config.Slide, out var singleSlide, out _) ? new List<PptSlideItem> { singleSlide } : slides);

            foreach (var slide in targetSlides)
            {
                var transition = slide.Doc.Descendants(PNs + "transition").FirstOrDefault();
                if (transition == null)
                {
                    return FailSpecialCondition(
                        $"Chưa áp dụng hiệu ứng chuyển tiếp (Transition) cho Slide {slide.Index}.",
                        fixAction: "Chọn thẻ Transitions -> Chọn hiệu ứng chuyển tiếp phù hợp (và bấm Apply To All nếu cần).");
                }

                // Kiểm tra Transition Type (e.g., Smoothly / Fade)
                if (!string.IsNullOrWhiteSpace(config.ExpectedTransition))
                {
                    var expNorm = config.ExpectedTransition.Trim().ToLowerInvariant();
                    bool transMatch = false;

                    // Smoothly trong PPT là biến thể của Fade không qua màu đen
                    if (expNorm.Contains("smooth") || expNorm.Contains("fade"))
                    {
                        var fade = transition.Element(PNs + "fade");
                        if (fade != null && fade.Attribute("thruBlk")?.Value != "1" && fade.Attribute("thruBlk")?.Value != "true")
                        {
                            transMatch = true;
                        }
                    }

                    if (!transMatch)
                    {
                        // Kiểm tra tên element con trong transition
                        var firstChild = transition.Elements().FirstOrDefault();
                        var name = firstChild?.Name.LocalName?.ToLowerInvariant();
                        if (name != null && (name.Contains(expNorm) || expNorm.Contains(name)))
                        {
                            transMatch = true;
                        }
                    }

                    if (!transMatch)
                    {
                        return FailSpecialCondition(
                            $"Hiệu ứng chuyển tiếp trên Slide {slide.Index} chưa đúng '{config.ExpectedTransition}'.",
                            fixAction: $"Chọn thẻ Transitions -> Transitions to This Slide -> Effect Options -> Chọn {config.ExpectedTransition}.");
                    }
                }

                // Kiểm tra Duration (P04-T4)
                if (config.ExpectedDurationSeconds.HasValue)
                {
                    var durStr = transition.Attribute(P14Ns + "dur")?.Value
                        ?? transition.Attribute(XNamespace.Get("http://schemas.microsoft.com/office/powerpoint/2010/main") + "dur")?.Value
                        ?? transition.Attribute("spd")?.Value;

                    long expectedMs = (long)(config.ExpectedDurationSeconds.Value * 1000); // 750 ms
                    bool durMatch = false;

                    if (long.TryParse(durStr, out var actualMs))
                    {
                        if (Math.Abs(actualMs - expectedMs) <= 100) durMatch = true;
                    }
                    else if (durStr != null)
                    {
                        durMatch = true; // Chấp nhận nếu có thuộc tính speed
                    }

                    if (!durMatch)
                    {
                        return FailSpecialCondition(
                            $"Thời lượng chuyển tiếp (Duration) trên Slide {slide.Index} chưa đúng {config.ExpectedDurationSeconds.Value}s.",
                            fixAction: $"Trên thẻ Transitions -> Timing -> Đặt Duration thành {config.ExpectedDurationSeconds.Value} (và chọn Apply To All).");
                    }
                }
            }

            return PassSpecialCondition("Đã áp dụng hiệu ứng chuyển tiếp thành công.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptAnimation(PptAnimationConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptAnimationConfig.");

            if (!TryResolveTargetSlide(package, config.Slide, out var slide, out var err))
                return FailSpecialCondition(err ?? "Không tìm thấy slide mục tiêu.");

            var timing = slide.Doc.Descendants(PNs + "timing").FirstOrDefault();
            if (timing == null)
            {
                return FailSpecialCondition(
                    $"Chưa có hiệu ứng hoạt họa (Animation) nào được thiết lập trên Slide {slide.Index}.",
                    fixAction: "Chọn đối tượng -> Animations -> Chọn hiệu ứng hoạt họa phù hợp.");
            }

            // P04-T3: Fly In (presetClass="entr", presetID="2")
            if (config.ExpectedEffect?.Contains("Fly In", StringComparison.OrdinalIgnoreCase) == true)
            {
                var hasFlyIn = timing.Descendants().Any(e =>
                {
                    var presetClass = e.Attribute("presetClass")?.Value;
                    var presetId = e.Attribute("presetID")?.Value;
                    return presetClass == "entr" && (presetId == "2" || presetId == "fly");
                });

                if (!hasFlyIn)
                {
                    return FailSpecialCondition(
                        $"Danh sách trên Slide {slide.Index} chưa được gán hiệu ứng Fly In.",
                        fixAction: "Chọn danh sách -> Animations -> Chọn Fly In.");
                }
            }

            // P05-T5: Motion Path Square
            if (config.MotionPathType?.Contains("Square", StringComparison.OrdinalIgnoreCase) == true)
            {
                var hasPath = timing.Descendants().Any(e =>
                {
                    var presetClass = e.Attribute("presetClass")?.Value;
                    var path = e.Descendants().Any(c => c.Name.LocalName.Contains("animMotion") || c.Name.LocalName.Contains("mpath"));
                    return presetClass == "path" || path;
                });

                if (!hasPath)
                {
                    return FailSpecialCondition(
                        $"Chưa thiết lập Motion Path Square cho hình ảnh MOS Certificate trên Slide {slide.Index}.",
                        fixAction: "Chọn hình ảnh -> Animations -> Motion Paths -> Shapes -> Effect Options -> Chọn Square.");
                }
            }

            return PassSpecialCondition($"Đã thiết lập hiệu ứng hoạt họa thành công trên Slide {slide.Index}.");
        }
    }
}
