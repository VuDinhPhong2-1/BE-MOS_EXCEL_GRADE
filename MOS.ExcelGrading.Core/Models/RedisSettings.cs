namespace MOS.ExcelGrading.Core.Models
{
    public class RedisSettings
    {
        public bool Enabled { get; set; }
        public string? ConnectionString { get; set; }
        public string InstanceName { get; set; } = "mos-grader";
        public int SchoolsTtlSeconds { get; set; } = 60;
        public int ClassesBySchoolTtlSeconds { get; set; } = 60;
        public int TeachersTtlSeconds { get; set; } = 60;
    }
}
