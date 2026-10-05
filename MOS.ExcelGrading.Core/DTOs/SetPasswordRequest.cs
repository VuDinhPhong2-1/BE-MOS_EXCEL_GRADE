using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.DTOs
{
    public class SetPasswordRequest
    {
        [Required(ErrorMessage = "Mật khẩu mới là bắt buộc")]
        [StringLength(100, MinimumLength = 6, ErrorMessage = "Mật khẩu phải từ 6 đến 100 ký tự")]
        public string NewPassword { get; set; } = string.Empty;
    }
}
