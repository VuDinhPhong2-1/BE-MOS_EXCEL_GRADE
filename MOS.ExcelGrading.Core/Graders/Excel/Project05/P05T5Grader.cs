using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Reflection;
using System.Xml;

namespace MOS.ExcelGrading.Core.Graders.Project05
{
    public class P05T5Grader : ITaskGrader
    {
        public string TaskId => "P05-T5";
        public string TaskName => "Đặt độ xoay hình ảnh trên Works về 0 độ";
        public decimal MaxScore => 4;

        public TaskResult Grade(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var ws = P05GraderHelpers.GetSheet(studentSheet, "Works");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Works'.");
                    return result;
                }

                var pictures = ws.Drawings.OfType<ExcelPicture>().ToList();
                if (pictures.Count == 0)
                {
                    result.Errors.Add("Không tìm thấy hình ảnh trên sheet Works.");
                    return result;
                }

                decimal score = 1m; // Tìm thấy hình ảnh cần chấm.
                var okCount = 0;

                foreach (var pic in pictures)
                {
                    if (TryGetRotationDegrees(pic, out var rotationDegrees) &&
                        Math.Abs(rotationDegrees) <= 0.01d)
                    {
                        okCount++;
                    }
                    else
                    {
                        var shownValue = TryGetRotationDegrees(pic, out var actualDeg)
                            ? $"{actualDeg:0.##}"
                            : "khong-doc-duoc";
                        result.Errors.Add($"Hình '{pic.Name}' đang có rotation = {shownValue} độ (căn = 0 độ).");
                    }
                }

                score += Math.Round(3m * okCount / pictures.Count, 2);
                if (okCount == pictures.Count)
                {
                    result.Details.Add("Tất cả hình ảnh trên Works đã đặt rotation = 0 độ.");
                }
                else
                {
                    result.Details.Add($"Hình đúng rotation 0 độ: {okCount}/{pictures.Count}.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }

        private static bool TryGetRotationDegrees(ExcelPicture picture, out double degrees)
        {
            degrees = 0;
            var topNodeProp = picture.GetType().GetProperty(
                "TopNode",
                BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            var topNode = topNodeProp?.GetValue(picture) as XmlNode;
            var xfrm = topNode?.SelectSingleNode(".//*[local-name()='xfrm']") as XmlElement;
            if (xfrm == null)
            {
                return false;
            }

            var rawRot = xfrm.GetAttribute("rot");
            if (string.IsNullOrWhiteSpace(rawRot))
            {
                degrees = 0;
                return true;
            }

            if (!double.TryParse(rawRot, out var rotIn60000))
            {
                return false;
            }

            degrees = rotIn60000 / 60000d;
            return true;
        }
    }
}

// minor-sync: non-functional graders update
