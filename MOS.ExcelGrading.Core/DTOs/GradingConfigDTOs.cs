using System.Text.Json.Nodes;

namespace MOS.ExcelGrading.Core.DTOs
{
    public class GradingConfigListItemDto
    {
        public string Id { get; set; } = string.Empty;
        public string GradingApiEndpoint { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public int ProjectNumber { get; set; }
        public string ProjectCode { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public int Version { get; set; }
        public decimal NormalizedMaxScore { get; set; }
        public int TaskCount { get; set; }
        public int EnabledTaskCount { get; set; }
        public decimal RawMaxScore { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime? UpdatedAt { get; set; }
        public DateTime? PublishedAt { get; set; }
        public string? Summary { get; set; }
    }

    public class GradingConfigDetailDto : GradingConfigListItemDto
    {
        public List<GradingConfigTaskDto> Tasks { get; set; } = new();
    }

    public class GradingConfigTaskDto
    {
        public string TaskId { get; set; } = string.Empty;
        public string TaskName { get; set; } = string.Empty;
        public decimal MaxScore { get; set; }
        public bool Enabled { get; set; } = true;
        public int SortOrder { get; set; }
        public string RuleType { get; set; } = string.Empty;
        public JsonNode? Rule { get; set; }
    }

    public class ImportGradingConfigFromCodeRequest
    {
        public string GradingApiEndpoint { get; set; } = string.Empty;
        public string? DisplayName { get; set; }
        public bool Publish { get; set; }
        public string? Summary { get; set; }
    }

    public class UpdateGradingConfigRequest
    {
        public string? DisplayName { get; set; }
        public string? Summary { get; set; }
        public List<UpdateGradingConfigTaskRequest> Tasks { get; set; } = new();
    }

    public class UpdateGradingConfigTaskRequest
    {
        public string TaskId { get; set; } = string.Empty;
        public string TaskName { get; set; } = string.Empty;
        public decimal MaxScore { get; set; }
        public bool Enabled { get; set; } = true;
        public int SortOrder { get; set; }
        public string RuleType { get; set; } = string.Empty;
        public JsonNode? Rule { get; set; }
    }

    public class PublishGradingConfigRequest
    {
        public string? Summary { get; set; }
    }

    public class RestoreGradingConfigVersionRequest
    {
        public string? Summary { get; set; }
    }

    public class TestGradingConfigRequest
    {
        public GradingConfigDetailDto? ConfigOverride { get; set; }
    }

    public class GradingConfigTestRunDto
    {
        public string Id { get; set; } = string.Empty;
        public string GradingConfigId { get; set; } = string.Empty;
        public string GradingApiEndpoint { get; set; } = string.Empty;
        public int ConfigVersion { get; set; }
        public bool UsedOverride { get; set; }
        public string FileName { get; set; } = string.Empty;
        public decimal TotalScore { get; set; }
        public decimal MaxScore { get; set; }
        public double Percentage { get; set; }
        public string Status { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; }
        public string? CreatedBy { get; set; }
        public string? Error { get; set; }
    }

    public class GradingConfigVersionDto
    {
        public string Id { get; set; } = string.Empty;
        public string GradingConfigId { get; set; } = string.Empty;
        public int Version { get; set; }
        public string Action { get; set; } = string.Empty;
        public string? Summary { get; set; }
        public DateTime CreatedAt { get; set; }
        public string? CreatedBy { get; set; }
    }

    public class GradingRuleTypeDto
    {
        public string RuleType { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
    }
}