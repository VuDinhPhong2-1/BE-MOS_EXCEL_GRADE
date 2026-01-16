// MOS.ExcelGrading.Core/DTOs/AssignmentDTOs.cs
namespace MOS.ExcelGrading.Core.DTOs
{
    public class CreateAssignmentRequest
    {
        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
        public string ClassId { get; set; } = string.Empty;
        public double MaxScore { get; set; } = 10;

        // ========== THÊM MỚI ==========
        /// <summary>
        /// Loại chấm điểm: "auto" hoặc "manual"
        /// </summary>
        public string GradingType { get; set; } = "manual";

        /// <summary>
        /// API endpoint để chấm điểm (chỉ dùng khi GradingType = "auto")
        /// Ví dụ: "project09", "project10"
        /// </summary>
        public string? GradingApiEndpoint { get; set; }
    }

    public class UpdateAssignmentRequest
    {
        public string? Name { get; set; }
        public string? Description { get; set; }
        public double? MaxScore { get; set; }
        public bool? IsActive { get; set; }

        // ========== THÊM MỚI ==========
        public string? GradingType { get; set; }
        public string? GradingApiEndpoint { get; set; }
    }

    public class AssignmentResponse
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
        public string ClassId { get; set; } = string.Empty;
        public double MaxScore { get; set; }
        public DateTime CreatedAt { get; set; }
        public bool IsActive { get; set; }
        public string? CreatedBy { get; set; }
        public string? CreatedByName { get; set; }
        public DateTime? UpdatedAt { get; set; }

        // ========== THÊM MỚI ==========
        public string GradingType { get; set; } = "manual";
        public string? GradingApiEndpoint { get; set; }
    }

    public class AssignmentWithStatsResponse : AssignmentResponse
    {
        public int TotalStudents { get; set; }
        public int GradedStudents { get; set; }
        public double AverageScore { get; set; }
        public double CompletionRate { get; set; }
    }

    // ========== THÊM MỚI: DTO CHO DANH SÁCH GRADING ENDPOINTS ==========
    public class GradingEndpointInfo
    {
        public string Endpoint { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public int MaxScore { get; set; } // ----> Sử dụng kiểu int
    }

}
