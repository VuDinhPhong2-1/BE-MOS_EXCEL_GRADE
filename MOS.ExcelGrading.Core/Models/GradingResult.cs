namespace MOS.ExcelGrading.Core.Models
{
    public class GradingResult
    {
        public string ProjectId { get; set; } = string.Empty;
        public string ProjectName { get; set; } = string.Empty;
        public decimal TotalScore { get; set; }
        public decimal MaxScore { get; set; }
        public double Percentage => MaxScore > 0 ? (double)(TotalScore / MaxScore * 100) : 0;
        public List<TaskResult> TaskResults { get; set; } = new();
        public DateTime GradedAt { get; set; } = DateTime.UtcNow;
        public string Status => Percentage >= 80 ? "Excellent" :
                               Percentage >= 60 ? "Good" :
                               Percentage >= 40 ? "Fair" : "Poor";
    }

    public class TaskResult
    {
        public string TaskId { get; set; } = string.Empty;
        public string TaskName { get; set; } = string.Empty;
        public decimal Score { get; set; }
        public decimal MaxScore { get; set; }
        public bool IsPassed => Score >= MaxScore * 0.5m;
        public List<string> Details { get; set; } = new();
        public List<string> Errors { get; set; } = new();
    }
}
