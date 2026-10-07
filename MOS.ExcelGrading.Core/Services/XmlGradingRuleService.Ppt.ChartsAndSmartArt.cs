using System.Xml.Linq;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public partial class XmlGradingRuleService
    {
        private static SpecialConditionEvalOutcome EvaluatePptChartLegend(PptChartLegendConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptChartLegendConfig.");

            // Tìm chart parts trong bài thuyết trình
            var chartParts = package.XmlParts.Keys
                .Where(k => k.StartsWith("ppt/charts/chart", StringComparison.OrdinalIgnoreCase) && k.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (chartParts.Count == 0)
            {
                return FailSpecialCondition("Không tìm thấy biểu đồ nào trong tệp PowerPoint.");
            }

            var expectedPos = string.IsNullOrWhiteSpace(config.ExpectedPosition) ? "t" : config.ExpectedPosition.Trim().ToLowerInvariant();
            bool matched = false;

            foreach (var cp in chartParts)
            {
                if (!package.XmlParts.TryGetValue(cp, out var xml)) continue;
                try
                {
                    var doc = XDocument.Parse(xml);
                    var legend = doc.Descendants(CNs + "legend").FirstOrDefault();
                    if (legend != null)
                    {
                        var legendPos = legend.Descendants(CNs + "legendPos").FirstOrDefault()?.Attribute("val")?.Value?.ToLowerInvariant();
                        if (legendPos == expectedPos)
                        {
                            matched = true;
                            break;
                        }
                    }
                }
                catch
                {
                }
            }

            if (!matched)
            {
                return FailSpecialCondition(
                    "Chú giải biểu đồ (Legend) chưa được đặt ở phía trên (Top) biểu đồ.",
                    fixAction: "Chọn biểu đồ -> Chart Tools Design -> Add Chart Element -> Legend -> Top.");
            }

            return PassSpecialCondition("Đã đặt vị trí chú giải biểu đồ ở phía trên thành công.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptChartType(PptChartTypeConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptChartTypeConfig.");

            // Tìm chart parts (kể cả chartEx mở rộng của Office 2016+ như Pareto / Histogram)
            var chartExParts = package.XmlParts.Keys
                .Where(k => k.StartsWith("ppt/charts/chartEx", StringComparison.OrdinalIgnoreCase))
                .ToList();

            var standardChartParts = package.XmlParts.Keys
                .Where(k => k.StartsWith("ppt/charts/chart", StringComparison.OrdinalIgnoreCase))
                .ToList();

            bool isPareto = false;
            // 1. Kiểm tra trong chartEx (Pareto Histogram trong PPT lưu ở dạng cx:chart)
            foreach (var cp in chartExParts)
            {
                if (!package.XmlParts.TryGetValue(cp, out var xml)) continue;
                if (xml.Contains("paretoLine", StringComparison.OrdinalIgnoreCase)
                    || xml.Contains("pareto", StringComparison.OrdinalIgnoreCase)
                    || xml.Contains("histogram", StringComparison.OrdinalIgnoreCase))
                {
                    isPareto = true;
                    break;
                }
            }

            // 2. Kiểm tra trong standard charts nếu không tìm thấy chartEx
            if (!isPareto)
            {
                foreach (var cp in standardChartParts)
                {
                    if (!package.XmlParts.TryGetValue(cp, out var xml)) continue;
                    if (xml.Contains("pareto", StringComparison.OrdinalIgnoreCase)
                        || xml.Contains("histogram", StringComparison.OrdinalIgnoreCase))
                    {
                        isPareto = true;
                        break;
                    }
                }
            }

            if (!isPareto)
            {
                return FailSpecialCondition(
                    $"Chưa chèn đúng biểu đồ {config.ExpectedChartType} trên Slide 5.",
                    fixAction: "Chọn Slide 5 -> Insert -> Chart -> Histogram -> Chọn Pareto -> OK.");
            }

            return PassSpecialCondition($"Đã chèn biểu đồ {config.ExpectedChartType} thành công trên Slide 5.");
        }

        private static SpecialConditionEvalOutcome EvaluatePptSmartArt(PptSmartArtConfig? config, OfficePackage package)
        {
            if (config == null) return FailSpecialCondition("Chưa cấu hình pptSmartArtConfig.");

            var diagramDataParts = package.XmlParts.Keys
                .Where(k => k.StartsWith("ppt/diagrams/data", StringComparison.OrdinalIgnoreCase))
                .ToList();

            var diagramLayoutParts = package.XmlParts.Keys
                .Where(k => k.StartsWith("ppt/diagrams/layout", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (diagramDataParts.Count == 0 && diagramLayoutParts.Count == 0)
            {
                return FailSpecialCondition("Không tìm thấy đồ họa SmartArt nào trong bài thuyết trình.");
            }

            // 1. Kiểm tra Layout Name / Unique ID (Target List hoặc Upward Arrow)
            if (!string.IsNullOrWhiteSpace(config.ExpectedLayoutName) || !string.IsNullOrWhiteSpace(config.ExpectedLayoutIdEnding))
            {
                var targetExp = (config.ExpectedLayoutName ?? config.ExpectedLayoutIdEnding)!.Trim();
                bool layoutMatched = false;

                // Ghép toàn bộ nội dung diagram để tìm kiếm định danh layout
                var combinedDiagrams = string.Join(" ", package.XmlParts
                    .Where(p => p.Key.StartsWith("ppt/diagrams/", StringComparison.OrdinalIgnoreCase))
                    .Select(p => p.Value));

                if (targetExp.Contains("Target", StringComparison.OrdinalIgnoreCase))
                {
                    if (combinedDiagrams.Contains("target", StringComparison.OrdinalIgnoreCase) || combinedDiagrams.Contains("Target List", StringComparison.OrdinalIgnoreCase))
                        layoutMatched = true;
                }
                else if (targetExp.Contains("Upward", StringComparison.OrdinalIgnoreCase) || targetExp.Contains("Arrow", StringComparison.OrdinalIgnoreCase))
                {
                    if (combinedDiagrams.Contains("upward", StringComparison.OrdinalIgnoreCase) || combinedDiagrams.Contains("arrow", StringComparison.OrdinalIgnoreCase))
                        layoutMatched = true;
                }
                else
                {
                    if (combinedDiagrams.Contains(targetExp, StringComparison.OrdinalIgnoreCase))
                        layoutMatched = true;
                }

                if (!layoutMatched)
                {
                    return FailSpecialCondition(
                        $"Đồ họa SmartArt chưa đúng mẫu '{targetExp}'.",
                        fixAction: $"Chọn SmartArt -> SmartArt Tools Design -> Layouts -> Chọn {targetExp}.");
                }
            }

            // 2. Kiểm tra node texts (P05-T4: Enroll Students)
            if (config.ExpectedNodeTexts.Count > 0)
            {
                var combinedDataText = string.Join(" ", diagramDataParts
                    .Where(p => package.XmlParts.ContainsKey(p))
                    .Select(p => package.XmlParts[p]));

                var normData = NormalizePlainText(combinedDataText);

                foreach (var expNode in config.ExpectedNodeTexts)
                {
                    var normNode = NormalizePlainText(expNode);
                    if (!normData.Contains(normNode, StringComparison.OrdinalIgnoreCase))
                    {
                        return FailSpecialCondition(
                            $"Đồ họa SmartArt chưa chứa hình dạng có văn bản '{expNode}'.",
                            fixAction: $"Chọn hình SmartArt -> Add Shape -> Nhập văn bản '{expNode}'.");
                    }
                }
            }

            return PassSpecialCondition("Đã cấu hình SmartArt thành công.");
        }
    }
}
