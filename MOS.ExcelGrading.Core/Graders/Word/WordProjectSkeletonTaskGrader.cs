using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word
{
    /// <summary>
    /// Skeleton grader cho Word project. Dùng để mở kiến trúc trước khi implement rule thực tế.
    /// </summary>
    public class WordProjectSkeletonTaskGrader : IWordTaskGrader
    {
        private readonly int _projectNumber;

        public WordProjectSkeletonTaskGrader(int projectNumber)
        {
            _projectNumber = projectNumber;
        }

        public string TaskId => "W-T01";
        public string TaskName => $"Word Project {_projectNumber:00} - Skeleton validation";
        public decimal MaxScore => 1m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore,
                Score = 0m
            };

            if (!studentDocument.HasMainDocumentPart)
            {
                result.Errors.Add("Thiếu part bắt buộc: word/document.xml.");
                return result;
            }

            result.Details.Add(
                $"Nhận diện thành công file Word project {_projectNumber:00}. " +
                $"Số part trong package: {studentDocument.PartCount}.");
            result.Details.Add("Skeleton mode: chưa có rule chấm chi tiết cho project này.");
            return result;
        }
    }
}
