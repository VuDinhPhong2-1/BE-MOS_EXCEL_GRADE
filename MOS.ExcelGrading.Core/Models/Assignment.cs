// MOS.ExcelGrading.Core/Models/Assignment.cs
using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

namespace MOS.ExcelGrading.Core.Models
{
    public class Assignment
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = ObjectId.GenerateNewId().ToString();

        [Required(ErrorMessage = "Tên bài tập là bắt buộc")]
        [StringLength(200)]
        [BsonElement("name")]
        public string Name { get; set; } = string.Empty;

        [StringLength(1000)]
        [BsonElement("description")]
        public string? Description { get; set; }

        [Required]
        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("classId")]
        public string ClassId { get; set; } = string.Empty;

        [Required]
        [Range(0, 1000)]
        [BsonElement("maxScore")]
        public double MaxScore { get; set; } = 10;


        // ========== THÊM MỚI: LIÊN KẾT VỚI GRADING API ==========
        /// <summary>
        /// API endpoint để chấm điểm (vd: "project09", "project10", "project11")
        /// </summary>
        [BsonElement("gradingApiEndpoint")]
        public string? GradingApiEndpoint { get; set; }

        /// <summary>
        /// Loại bài tập: "auto" (tự động chấm), "manual" (chấm thủ công)
        /// </summary>
        [BsonElement("gradingType")]
        public string GradingType { get; set; } = "manual"; // "auto" | "manual"

        [BsonElement("isActive")]
        public bool IsActive { get; set; } = true;

        // ========== METADATA ==========
        [BsonElement("createdAt")]
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("createdBy")]
        public string? CreatedBy { get; set; }

        [BsonElement("updatedAt")]
        public DateTime? UpdatedAt { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("updatedBy")]
        public string? UpdatedBy { get; set; }
    }

    // ========== ĐỊNH NGHĨA GRADING TYPES ==========
    public static class GradingTypes
    {
        public const string Auto = "auto";
        public const string Manual = "manual";
    }

    // ========== NHÓM MÔN CHẤM ==========
    public static class GradingApiSubjects
    {
        public const string Excel = "excel";
        public const string Word = "word";
        public const string Ppt = "ppt";
    }

    // ========== THANG ĐIỂM PRACTICE ==========
    public sealed class PracticeDefinition
    {
        public string Code { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public int TotalScore { get; init; }
        public int ProjectCount { get; init; }
        public decimal ScorePerProject =>
            ProjectCount > 0
                ? Math.Round((decimal)TotalScore / ProjectCount, 2, MidpointRounding.AwayFromZero)
                : 0m;
    }

    public static class PracticeScoring
    {
        public const int PracticeTotalScore = 1000;

        // Theo cấu hình yêu cầu hiện tại:
        // Practice 01: 8 bài | Practice 02: 9 bài | Practice 03: 17 bài
        private static readonly PracticeDefinition Practice01 = new()
        {
            Code = "practice01",
            Name = "Practice 01",
            TotalScore = PracticeTotalScore,
            ProjectCount = 8
        };

        private static readonly PracticeDefinition Practice02 = new()
        {
            Code = "practice02",
            Name = "Practice 02",
            TotalScore = PracticeTotalScore,
            ProjectCount = 9
        };

        private static readonly PracticeDefinition Practice03 = new()
        {
            Code = "practice03",
            Name = "Practice 03",
            TotalScore = PracticeTotalScore,
            ProjectCount = 17
        };

        public static PracticeDefinition ResolveByProjectNumber(int projectNumber)
        {
            if (projectNumber >= 1 && projectNumber <= 8)
            {
                return Practice01;
            }

            if (projectNumber >= 9 && projectNumber <= 16)
            {
                return Practice02;
            }

            if (projectNumber >= 17 && projectNumber <= 24)
            {
                return Practice03;
            }

            return Practice03;
        }

        public static decimal CalculateProjectMaxScore(int projectNumber)
        {
            return ResolveByProjectNumber(projectNumber).ScorePerProject;
        }
    }

    // ========== ĐỊNH NGHĨA GRADING API ENDPOINTS ==========
    public static class GradingApiEndpoints
    {
        public const string Project01 = $"{GradingApiSubjects.Excel}/project01";
        public const string Project02 = $"{GradingApiSubjects.Excel}/project02";
        public const string Project03 = $"{GradingApiSubjects.Excel}/project03";
        public const string Project04 = $"{GradingApiSubjects.Excel}/project04";
        public const string Project05 = $"{GradingApiSubjects.Excel}/project05";
        public const string Project06 = $"{GradingApiSubjects.Excel}/project06";
        public const string Project07 = $"{GradingApiSubjects.Excel}/project07";
        public const string Project08 = $"{GradingApiSubjects.Excel}/project08";
        public const string Project09 = $"{GradingApiSubjects.Excel}/project09";
        public const string Project10 = $"{GradingApiSubjects.Excel}/project10";
        public const string Project11 = $"{GradingApiSubjects.Excel}/project11";
        public const string Project12 = $"{GradingApiSubjects.Excel}/project12";
        public const string Project13 = $"{GradingApiSubjects.Excel}/project13";
        public const string Project14 = $"{GradingApiSubjects.Excel}/project14";
        public const string Project15 = $"{GradingApiSubjects.Excel}/project15";
        public const string Project16 = $"{GradingApiSubjects.Excel}/project16";
        public const string Project17 = $"{GradingApiSubjects.Excel}/project17";
        public const string Project18 = $"{GradingApiSubjects.Excel}/project18";
        public const string Project19 = $"{GradingApiSubjects.Excel}/project19";
        public const string Project20 = $"{GradingApiSubjects.Excel}/project20";
        public const string Project21 = $"{GradingApiSubjects.Excel}/project21";
        public const string Project22 = $"{GradingApiSubjects.Excel}/project22";
        public const string Project23 = $"{GradingApiSubjects.Excel}/project23";
        public const string Project24 = $"{GradingApiSubjects.Excel}/project24";

        private static readonly Regex ProjectRegex = new(
            @"^project(?<number>\d{1,2})$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex SubjectProjectRegex = new(
            @"^(?<subject>excel|word|ppt|powerpoint)/project(?<number>\d{1,2})$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static List<string> GetAllEndpoints() =>
            Enumerable.Range(1, 24)
                .Select(ToExcelProjectEndpoint)
                .ToList();

        public static string ToExcelProjectEndpoint(int projectNumber) =>
            $"{GradingApiSubjects.Excel}/project{projectNumber:00}";

        public static string NormalizeEndpoint(string? endpoint)
        {
            if (string.IsNullOrWhiteSpace(endpoint))
            {
                return string.Empty;
            }

            var normalized = endpoint.Trim().Replace("\\", "/", StringComparison.Ordinal).TrimStart('/');
            normalized = normalized.ToLowerInvariant();

            if (normalized.StartsWith("api/grading/", StringComparison.Ordinal))
            {
                normalized = normalized["api/grading/".Length..];
            }
            else if (normalized.StartsWith("grading/", StringComparison.Ordinal))
            {
                normalized = normalized["grading/".Length..];
            }

            var subjectMatch = SubjectProjectRegex.Match(normalized);
            if (subjectMatch.Success)
            {
                var subjectRaw = subjectMatch.Groups["subject"].Value.ToLowerInvariant();
                var subject = subjectRaw == "powerpoint" ? GradingApiSubjects.Ppt : subjectRaw;
                var projectNumber = int.Parse(subjectMatch.Groups["number"].Value);
                return $"{subject}/project{projectNumber:00}";
            }

            var projectMatch = ProjectRegex.Match(normalized);
            if (projectMatch.Success)
            {
                var projectNumber = int.Parse(projectMatch.Groups["number"].Value);
                return ToExcelProjectEndpoint(projectNumber);
            }

            return normalized;
        }

        public static bool IsValidEndpoint(string endpoint)
        {
            var normalized = NormalizeEndpoint(endpoint);
            return GetAllEndpoints().Contains(normalized);
        }

        public static bool TryExtractProjectNumber(string? endpoint, out int projectNumber)
        {
            projectNumber = 0;
            var normalized = NormalizeEndpoint(endpoint);
            var match = SubjectProjectRegex.Match(normalized);
            if (!match.Success)
            {
                return false;
            }

            return int.TryParse(match.Groups["number"].Value, out projectNumber);
        }
    }
}
