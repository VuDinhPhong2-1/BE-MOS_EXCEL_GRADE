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
        private readonly IAnalyticsService _analyticsService;
        private readonly ILogger<GradingController> _logger;

        public GradingController(
            IGradingService gradingService,
            IAnalyticsService analyticsService,
            ILogger<GradingController> logger)
        {
            _gradingService = gradingService;
            _analyticsService = analyticsService;
            _logger = logger;
        }

        /// <summary>
        /// Chấm điểm Project 01
        /// Chỉ Teacher và Admin mới được phép sử dụng
        /// </summary>
        [HttpPost("project01")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject01(
            [FromForm] IFormFile studentFile,
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                var hasPermission = User.Claims.Any(c =>
                    c.Type == "permission" && c.Value == Permissions.CreateGrades);

                if (!hasPermission)
                {
                    _logger.LogWarning($"User {username} (ID: {userId}) không có quyền {Permissions.CreateGrades}");
                    return Forbid();
                }

                if (studentFile == null)
                    return BadRequest(new { error = "Cần cung cấp file: studentFile" });

                if (!IsExcelFile(studentFile))
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject01Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project01,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading project01");
                return StatusCode(500, new { error = "Lỗi hệ thống khi chấm điểm" });
            }
        }

        /// <summary>
        /// Cham diem Project 02
        /// Chi Teacher va Admin duoc phep
        /// </summary>
        [HttpPost("project02")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject02(
            [FromForm] IFormFile studentFile,
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                var hasPermission = User.Claims.Any(c =>
                    c.Type == "permission" && c.Value == Permissions.CreateGrades);

                if (!hasPermission)
                {
                    _logger.LogWarning($"User {username} (ID: {userId}) khong co quyen {Permissions.CreateGrades}");
                    return Forbid();
                }

                if (studentFile == null)
                    return BadRequest(new { error = "Can cung cap file: studentFile" });

                if (!IsExcelFile(studentFile))
                    return BadRequest(new { error = "File phai co dinh dang .xlsx hoac .xlsm" });

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject02Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project02,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading project02");
                return StatusCode(500, new { error = "Loi he thong khi cham diem" });
            }
        }

        /// <summary>
        /// Cham diem Project 03
        /// Chi Teacher va Admin duoc phep
        /// </summary>
        [HttpPost("project03")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject03(
            [FromForm] IFormFile studentFile,
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                var hasPermission = User.Claims.Any(c =>
                    c.Type == "permission" && c.Value == Permissions.CreateGrades);

                if (!hasPermission)
                {
                    _logger.LogWarning($"User {username} (ID: {userId}) khong co quyen {Permissions.CreateGrades}");
                    return Forbid();
                }

                if (studentFile == null)
                    return BadRequest(new { error = "Can cung cap file: studentFile" });

                if (!IsExcelFile(studentFile))
                    return BadRequest(new { error = "File phai co dinh dang .xlsx hoac .xlsm" });

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject03Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project03,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading project03");
                return StatusCode(500, new { error = "Loi he thong khi cham diem" });
            }
        }

        /// <summary>
        /// Cham diem Project 04
        /// Chi Teacher va Admin duoc phep
        /// </summary>
        [HttpPost("project04")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject04(
            [FromForm] IFormFile studentFile,
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                var hasPermission = User.Claims.Any(c =>
                    c.Type == "permission" && c.Value == Permissions.CreateGrades);

                if (!hasPermission)
                {
                    _logger.LogWarning($"User {username} (ID: {userId}) khong co quyen {Permissions.CreateGrades}");
                    return Forbid();
                }

                if (studentFile == null)
                    return BadRequest(new { error = "Can cung cap file: studentFile" });

                if (!IsExcelFile(studentFile))
                    return BadRequest(new { error = "File phai co dinh dang .xlsx hoac .xlsm" });

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject04Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project04,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading project04");
                return StatusCode(500, new { error = "Loi he thong khi cham diem" });
            }
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
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
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

                if (studentFile == null)
                {
                    return BadRequest(new { error = "Cần cung cấp file: studentFile" });
                }

                // ✅ LOG THÔNG TIN NGƯỜI DÙNG VÀ FILE
                _logger.LogInformation(
                    $"[GRADING] User: {username} (ID: {userId}, Role: {userRole}) | " +
                    $"Student file: {studentFile.FileName} ({studentFile.Length:N0} bytes)");

                if (!IsExcelFile(studentFile))
                {
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });
                }

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject09Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project09,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

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

                return StatusCode(500, new { error = "Lỗi hệ thống khi chấm điểm" });
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
