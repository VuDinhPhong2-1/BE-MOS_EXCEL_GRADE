namespace MOS.ExcelGrading.Core.Models
{
    public static class TaskNameFormatter
    {
        private static readonly System.Text.RegularExpressions.Regex TaskNumberRegex = new(
            @"^[A-Z]+\d{2}-T(?<taskNumber>\d+)$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase |
            System.Text.RegularExpressions.RegexOptions.CultureInvariant |
            System.Text.RegularExpressions.RegexOptions.Compiled);

        private static readonly System.Text.RegularExpressions.Regex ExistingPrefixRegex = new(
            @"^\s*Task\s+\d+\s*:",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase |
            System.Text.RegularExpressions.RegexOptions.CultureInvariant |
            System.Text.RegularExpressions.RegexOptions.Compiled);

        public static string Format(string? taskId, string? taskName)
        {
            var normalizedTaskId = (taskId ?? string.Empty).Trim();
            var normalizedTaskName = (taskName ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(normalizedTaskName))
            {
                return normalizedTaskName;
            }

            if (ExistingPrefixRegex.IsMatch(normalizedTaskName))
            {
                return normalizedTaskName;
            }

            var matched = TaskNumberRegex.Match(normalizedTaskId);
            if (!matched.Success)
            {
                return normalizedTaskName;
            }

            var taskNumber = matched.Groups["taskNumber"].Value.TrimStart('0');
            if (string.IsNullOrWhiteSpace(taskNumber))
            {
                taskNumber = "0";
            }

            return $"Task {taskNumber}: {normalizedTaskName}";
        }

        public static string RemovePrefix(string? taskName)
        {
            var normalizedTaskName = (taskName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalizedTaskName))
            {
                return string.Empty;
            }

            return ExistingPrefixRegex.Replace(normalizedTaskName, string.Empty).Trim();
        }
    }

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
            if (string.IsNullOrWhiteSpace(item))
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
            }
        }
    }

    public class TaskDisplayIssue
    {
        public string Heading { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public string FixAction { get; set; } = string.Empty;
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
        private string _taskId = string.Empty;
        private string _taskName = string.Empty;

        public string TaskId
        {
            get => _taskId;
            set
            {
                _taskId = value ?? string.Empty;
                _taskName = TaskNameFormatter.Format(_taskId, _taskName);
            }
        }

        public string TaskName
        {
            get => _taskName;
            set => _taskName = TaskNameFormatter.Format(_taskId, value);
        }

        public decimal Score { get; set; }
        public decimal MaxScore { get; set; }
        public bool IsPassed => Score >= MaxScore * 0.5m;
        public List<string> Details { get; set; } = new();
        public List<string> Errors { get; set; } = new();
        public SingleFixActionList FixActions { get; set; } = new();
        public List<TaskDisplayIssue> DisplayIssues { get; set; } = new();
    }

    public static class TaskResultIssueHelper
    {
        public static void AddIssue(TaskResult result, string message, string? fixAction = null)
        {
            if (result == null)
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(message))
            {
                result.Errors.Add(message.Trim());
            }

            if (!string.IsNullOrWhiteSpace(fixAction))
            {
                result.FixActions.Add(fixAction.Trim());
            }
        }
    }
}
