namespace MOS.ExcelGrading.Core.DTOs
{
    public class SchoolResponse
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Code { get; set; } = string.Empty;
        public string? Address { get; set; }
        public string? PhoneNumber { get; set; }
        public string? Email { get; set; }
        public string? Website { get; set; }
        public string? Description { get; set; }
        public string? Logo { get; set; }
        public string OwnerId { get; set; } = string.Empty;
        public string? OwnerName { get; set; }
        public DateTime CreatedAt { get; set; }
        public bool IsActive { get; set; }
    }
}
