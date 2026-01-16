using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Security.Claims;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize] // ✅ YÊU CẦU ĐĂNG NHẬP CHO TẤT CẢ ENDPOINT
    public class GradingController : ControllerBase
    {
        private readonly IGradingService _gradingService;
        private readonly ILogger<GradingController> _logger;

        public GradingController(
            IGradingService gradingService,
            ILogger<GradingController> logger)
        {
            _gradingService = gradingService;
            _logger = logger;
        }

        /// <summary>
        /// Chấm điểm Project 09
        /// Chỉ Teacher và Admin mới được phép sử dụng
        /// </summary>
        [HttpPost("project09")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")] // ✅ CHỈ TEACHER VÀ ADMIN
        [RequestSizeLimit(524288000)] // 500MB
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject09(
            [FromForm] IFormFile studentFile,
            [FromForm] IFormFile answerFile)
        {
            try
            {
                // ✅ LẤY THÔNG TIN NGƯỜI DÙNG TỪ TOKEN
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? "unknown";

                // ✅ KIỂM TRA PERMISSION (tùy chọn - nếu muốn kiểm tra chi tiết hơn)
                var hasPermission = User.Claims.Any(c =>
                    c.Type == "permission" && c.Value == Permissions.CreateGrades);

                if (!hasPermission)
                {
                    _logger.LogWarning($"User {username} (ID: {userId}) không có quyền {Permissions.CreateGrades}");
                    return Forbid(); // 403 Forbidden
                }

                if (studentFile == null || answerFile == null)
                {
                    return BadRequest(new { error = "Cần cung cấp cả 2 file: studentFile và answerFile" });
                }

                // ✅ LOG THÔNG TIN NGƯỜI DÙNG VÀ FILE
                _logger.LogInformation(
                    $"[GRADING] User: {username} (ID: {userId}, Role: {userRole}) | " +
                    $"Student file: {studentFile.FileName} ({studentFile.Length:N0} bytes) | " +
                    $"Answer file: {answerFile.FileName} ({answerFile.Length:N0} bytes)");

                if (!IsExcelFile(studentFile) || !IsExcelFile(answerFile))
                {
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });
                }

                using var studentStream = studentFile.OpenReadStream();
                using var answerStream = answerFile.OpenReadStream();

                var result = await _gradingService.GradeProject09Async(studentStream, answerStream);

                // ✅ LOG KẾT QUẢ THÀNH CÔNG
                _logger.LogInformation(
                    $"[GRADING SUCCESS] User: {username} (ID: {userId}) | " +
                    $"Score: {result.TotalScore}/{result.MaxScore}");

                return Ok(result);
            }
            catch (Exception ex)
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                _logger.LogError(ex,
                    $"[GRADING ERROR] User: {username} (ID: {userId}) | Error: {ex.Message}");

                return StatusCode(500, new { error = $"Lỗi hệ thống: {ex.Message}" });
            }
        }

        /// <summary>
        /// Health check - Không cần xác thực
        /// </summary>
        [HttpGet("health")]
        [AllowAnonymous] // ✅ CHO PHÉP TRUY CẬP KHÔNG CẦN TOKEN
        public IActionResult HealthCheck()
        {
            return Ok(new
            {
                status = "OK",
                timestamp = DateTime.UtcNow,
                maxUploadSize = "500MB"
            });
        }

        private bool IsExcelFile(IFormFile file)
        {
            var extension = Path.GetExtension(file.FileName).ToLower();
            return extension == ".xlsx" || extension == ".xlsm";
        }
    }
}
