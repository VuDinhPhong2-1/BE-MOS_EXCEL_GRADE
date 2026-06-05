namespace MOS.ExcelGrading.Core.Models
{
    public class SingleFixActionList : List<string>
    {
        public SingleFixActionList()
        {
        }

        public SingleFixActionList(IEnumerable<string>? items)
        {
            AddRange(items ?? Enumerable.Empty<string>());
        }

        public new void Add(string item)
        {
            if (Count > 0 || string.IsNullOrWhiteSpace(item))
            {
                return;
            }

            base.Add(item);
        }

        public new void AddRange(IEnumerable<string> collection)
        {
            if (collection == null)
            {
                return;
            }

            foreach (var item in collection)
            {
                Add(item);
                if (Count > 0)
                {
                    break;
                }
            }
        }
    }

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
        public SingleFixActionList FixActions { get; set; } = new();
    }
}
