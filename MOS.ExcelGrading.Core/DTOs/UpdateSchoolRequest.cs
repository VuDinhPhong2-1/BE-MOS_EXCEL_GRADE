namespace MOS.ExcelGrading.Core.DTOs
{
    public class UpdateSchoolRequest
    {
        public string? Name { get; set; }
        public string? Code { get; set; }
        public string? Address { get; set; }
        public string? PhoneNumber { get; set; }
        public string? Email { get; set; }
        public string? Website { get; set; }
        public string? Description { get; set; }
        public string? Logo { get; set; }
        public bool? IsActive { get; set; }
    }
}
