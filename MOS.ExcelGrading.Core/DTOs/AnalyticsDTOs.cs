namespace MOS.ExcelGrading.Core.DTOs
{
    public class ClassAnalyticsOverviewResponse
    {
        public string ClassId { get; set; } = string.Empty;
        public int TotalAttempts { get; set; }
        public int TotalStudents { get; set; }
        public double AveragePercentage { get; set; }
        public double PassRate { get; set; }
        public double WarningRate { get; set; }
    }

    public class WeakTaskResponse
    {
        public string TaskId { get; set; } = string.Empty;
        public string TaskName { get; set; } = string.Empty;
        public int AttemptCount { get; set; }
        public int FailedCount { get; set; }
        public double FailedRate { get; set; }
    }

    public class ProjectPerformanceResponse
    {
        public string ProjectEndpoint { get; set; } = string.Empty;
        public int AttemptCount { get; set; }
        public double AveragePercentage { get; set; }
        public double PassRate { get; set; }
    }
}
