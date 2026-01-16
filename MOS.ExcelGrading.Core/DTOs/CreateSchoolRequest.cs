using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.DTOs
{
    public class CreateSchoolRequest
    {
        [Required]
        [StringLength(200, MinimumLength = 3)]
        public string Name { get; set; } = string.Empty;

        public string Code { get; set; } = string.Empty;

        public string? Address { get; set; }
        public string? PhoneNumber { get; set; }
        public string? Email { get; set; }
        public string? Website { get; set; }
        public string? Description { get; set; }
        public string? Logo { get; set; }
    }
}
