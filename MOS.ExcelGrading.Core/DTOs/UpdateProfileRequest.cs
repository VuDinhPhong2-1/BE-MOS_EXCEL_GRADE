using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.DTOs
{
    public class UpdateProfileRequest
    {
        [StringLength(120, ErrorMessage = "Họ và tên không được vượt quá 120 ký tự")]
        public string? FullName { get; set; }

        [StringLength(25, ErrorMessage = "Số điện thoại không được vượt quá 25 ký tự")]
        public string? PhoneNumber { get; set; }

        [StringLength(500, ErrorMessage = "Ảnh đại diện không được vượt quá 500 ký tự")]
        public string? Avatar { get; set; }
    }
}
