using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.DTOs
{
    public class RegisterRequest
    {
        [Required(ErrorMessage = "Email là bắt buộc")]
        [EmailAddress(ErrorMessage = "Email không hợp lệ")]
        public string Email { get; set; } = string.Empty;

        [Required(ErrorMessage = "Username là bắt buộc")]
        [StringLength(50, MinimumLength = 3, ErrorMessage = "Username phải từ 3-50 ký tự")]
        public string Username { get; set; } = string.Empty;

        [Required(ErrorMessage = "Password là bắt buộc")]
        [StringLength(100, MinimumLength = 6, ErrorMessage = "Password phải từ 6-100 ký tự")]
        public string Password { get; set; } = string.Empty;

        // Thông tin bổ sung (optional)
        public string? FullName { get; set; }
        public string? PhoneNumber { get; set; }

        // Role (mặc định là Student nếu không truyền)
        public string? Role { get; set; }
    }
}
