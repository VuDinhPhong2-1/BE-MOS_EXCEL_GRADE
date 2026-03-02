using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.DTOs
{
    public class GoogleLoginRequest
    {
        [Required(ErrorMessage = "IdToken là bắt buộc")]
        public string IdToken { get; set; } = string.Empty;
    }
}
