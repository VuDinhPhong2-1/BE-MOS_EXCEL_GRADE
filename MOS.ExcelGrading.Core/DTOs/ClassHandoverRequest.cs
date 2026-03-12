using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.DTOs
{
    public class ClassHandoverRequest
    {
        [Required]
        public string TeacherId { get; set; } = string.Empty;
    }
}
