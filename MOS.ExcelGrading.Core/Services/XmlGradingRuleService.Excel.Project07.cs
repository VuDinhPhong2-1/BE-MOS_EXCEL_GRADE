using MOS.ExcelGrading.Core.Models;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace MOS.ExcelGrading.Core.Services
{
    public partial class XmlGradingRuleService
    {
        private static SpecialConditionEvalOutcome EvaluateExcelSparkline(
            ExcelSparklineConfig? config,
            OfficePackage package)
        {
            if (config == null)
            {
                return FailSpecialCondition("Chưa cấu hình kiểm tra Sparkline (excelSparklineConfig trống).");
            }

            if (!TryResolveExcelWorksheet(package, config.WorksheetName, config.SourceFile, out var worksheetPath, out var worksheetError))
            {
                return FailSpecialCondition(worksheetError);
            }

            if (!package.TryGetXmlDocument(worksheetPath, out var worksheetDoc, out var docError) || worksheetDoc.Root == null)
            {
                return FailSpecialCondition(docError ?? $"Không tìm thấy {worksheetPath} trong file học sinh.");
            }

            // Find all sparklineGroup elements across any namespace
            var sparklineGroups = worksheetDoc.Descendants()
                .Where(el => el.Name.LocalName.Equals("sparklineGroup", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (sparklineGroups.Count == 0)
            {
                return FailSpecialCondition(string.IsNullOrWhiteSpace(config.WorksheetName)
                    ? "Không tìm thấy biểu đồ Sparkline nào trong trang tính."
                    : $"Không tìm thấy biểu đồ Sparkline nào trên trang tính '{config.WorksheetName}'.");
            }

            var expectedType = config.SparklineType?.Trim().ToLowerInvariant() ?? "line";
            var expectedLocation = NormalizeExcelCellAddress(config.LocationRange);
            var expectedDataRange = NormalizeExcelFormulaReference(config.DataRange);

            foreach (var group in sparklineGroups)
            {
                var actualType = group.Attribute("type")?.Value?.ToLowerInvariant() ?? "line";
                if (!string.Equals(actualType, expectedType, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var sparklines = group.Descendants().Where(el => el.Name.LocalName.Equals("sparkline", StringComparison.OrdinalIgnoreCase));
                foreach (var sp in sparklines)
                {
                    var formulaElem = sp.Descendants().FirstOrDefault(el => el.Name.LocalName.Equals("f", StringComparison.OrdinalIgnoreCase));
                    var sqrefElem = sp.Descendants().FirstOrDefault(el => el.Name.LocalName.Equals("sqref", StringComparison.OrdinalIgnoreCase));

                    var actualFormula = NormalizeExcelFormulaReference(formulaElem?.Value);
                    var actualSqref = NormalizeExcelCellAddress(sqrefElem?.Value);

                    var locationMatched = string.IsNullOrWhiteSpace(expectedLocation) ||
                        string.Equals(actualSqref, expectedLocation, StringComparison.OrdinalIgnoreCase);

                    var dataRangeMatched = string.IsNullOrWhiteSpace(expectedDataRange) ||
                        ExcelDefinedNameRangeMatches(actualFormula, expectedDataRange);

                    if (locationMatched && dataRangeMatched)
                    {
                        var wsMsg = string.IsNullOrWhiteSpace(config.WorksheetName) ? "" : $" trên trang tính '{config.WorksheetName}'";
                        return PassSpecialCondition($"Đã tạo đúng biểu đồ Sparkline loại '{actualType}' tại ô {actualSqref}{wsMsg}.");
                    }
                }
            }

            var wsErr = string.IsNullOrWhiteSpace(config.WorksheetName) ? "" : $" trên trang tính '{config.WorksheetName}'";
            var typeErr = string.IsNullOrWhiteSpace(config.SparklineType) ? "" : $" loại {config.SparklineType}";
            var locErr = string.IsNullOrWhiteSpace(config.LocationRange) ? "" : $" tại ô {config.LocationRange}";
            var dataErr = string.IsNullOrWhiteSpace(config.DataRange) ? "" : $" với vùng dữ liệu {config.DataRange}";

            return FailSpecialCondition($"Chưa cấu hình đúng biểu đồ Sparkline{typeErr}{locErr}{dataErr}{wsErr}.");
        }

        private static void ValidateExcelSparklineSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelSparklineConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelSparklineConfig khong duoc null.");
                return;
            }

            ValidateExcelWorksheetLocator(config.WorksheetName, config.SourceFile, $"{taskPrefix}.specialCondition.excelSparklineConfig", result);

            if (string.IsNullOrWhiteSpace(config.LocationRange))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelSparklineConfig.locationRange khong duoc de trong.");
            }

            if (string.IsNullOrWhiteSpace(config.DataRange))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelSparklineConfig.dataRange khong duoc de trong.");
            }
        }

        private static SpecialConditionEvalOutcome EvaluateExcelTableRowDelete(
            ExcelTableRowDeleteConfig? config,
            OfficePackage package)
        {
            if (config == null || string.IsNullOrWhiteSpace(config.DeletedText))
            {
                return FailSpecialCondition("Chưa cấu hình nội dung cần xóa (excelTableRowDeleteConfig.deletedText trống).");
            }

            if (!TryResolveExcelWorksheet(package, config.WorksheetName, config.SourceFile, out var worksheetPath, out var worksheetError))
            {
                return FailSpecialCondition(worksheetError);
            }

            if (!package.TryGetXmlDocument(worksheetPath, out var worksheetDocument, out var worksheetDocumentError))
            {
                return FailSpecialCondition(worksheetDocumentError ?? $"Không tìm thấy {worksheetPath} trong file học sinh.");
            }

            var sharedStrings = GetExcelSharedStrings(package);
            var deletedText = config.DeletedText.Trim();
            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            var allCells = worksheetDocument.Descendants(x + "c").ToList();
            foreach (var cell in allCells)
            {
                var text = NormalizePlainText(GetExcelCellDisplayText(cell, sharedStrings));
                if (string.IsNullOrWhiteSpace(text)) continue;

                bool match = config.MatchWholeWord == false
                    ? text.Contains(deletedText, StringComparison.OrdinalIgnoreCase)
                    : string.Equals(text, deletedText, StringComparison.OrdinalIgnoreCase);

                if (match)
                {
                    var cellAddr = cell.Attribute("r")?.Value ?? "chưa xác định";
                    var wsDisplay = string.IsNullOrWhiteSpace(config.WorksheetName) ? "trang tính" : $"trang tính '{config.WorksheetName}'";
                    return FailSpecialCondition($"Vẫn còn dữ liệu '{config.DeletedText}' tại ô {cellAddr} trên {wsDisplay}.");
                }
            }

            // Kiểm tra thêm các bảng (table*.xml) liên kết nếu có
            var tableParts = ResolveExcelTableParts(package, config.WorksheetName, null);
            foreach (var tablePath in tableParts)
            {
                if (package.TryGetXmlDocument(tablePath, out var tableDoc, out _) && tableDoc.Root != null)
                {
                    var rawXml = tableDoc.Root.ToString();
                    if (rawXml.Contains(deletedText, StringComparison.OrdinalIgnoreCase))
                    {
                        var wsDisplay = string.IsNullOrWhiteSpace(config.WorksheetName) ? "trang tính" : $"trang tính '{config.WorksheetName}'";
                        return FailSpecialCondition($"Dữ liệu '{config.DeletedText}' vẫn còn xuất hiện trong cấu trúc bảng trên {wsDisplay}.");
                    }
                }
            }

            var successWs = string.IsNullOrWhiteSpace(config.WorksheetName) ? "trang tính" : $"trang tính '{config.WorksheetName}'";
            return PassSpecialCondition($"Đã xóa thành công hàng chứa '{config.DeletedText}' trên {successWs}.");
        }

        private static void ValidateExcelTableRowDeleteSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelTableRowDeleteConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelTableRowDeleteConfig khong duoc null.");
                return;
            }

            ValidateExcelWorksheetLocator(config.WorksheetName, config.SourceFile, $"{taskPrefix}.specialCondition.excelTableRowDeleteConfig", result);

            if (string.IsNullOrWhiteSpace(config.DeletedText))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelTableRowDeleteConfig.deletedText khong duoc de trong.");
            }
        }
    }
}
