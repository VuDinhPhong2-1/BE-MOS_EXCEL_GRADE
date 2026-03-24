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
        [HttpPost("excel/project01")]
        [HttpPost("project01")] // Legacy alias
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

                if (!HasCreateGradesPermission(userId, username))
                {
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
        /// Chấm điểm dự án 02
        /// Chi Teacher va Admin duoc phep
        /// </summary>
        [HttpPost("excel/project02")]
        [HttpPost("project02")] // Legacy alias
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

                if (!HasCreateGradesPermission(userId, username))
                {
                    return Forbid();
                }

                if (studentFile == null)
                    return BadRequest(new { error = "Cần cung cấp file: studentFile" });

                if (!IsExcelFile(studentFile))
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });

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
                return StatusCode(500, new { error = "Lỗi hệ thống khi chấm điểm" });
            }
        }

        /// <summary>
        /// Chấm điểm dự án 03
        /// Chi Teacher va Admin duoc phep
        /// </summary>
        [HttpPost("excel/project03")]
        [HttpPost("project03")] // Legacy alias
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

                if (!HasCreateGradesPermission(userId, username))
                {
                    return Forbid();
                }

                if (studentFile == null)
                    return BadRequest(new { error = "Cần cung cấp file: studentFile" });

                if (!IsExcelFile(studentFile))
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });

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
                return StatusCode(500, new { error = "Lỗi hệ thống khi chấm điểm" });
            }
        }

        /// <summary>
        /// Chấm điểm dự án 04
        /// Chi Teacher va Admin duoc phep
        /// </summary>
        [HttpPost("excel/project04")]
        [HttpPost("project04")] // Legacy alias
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

                if (!HasCreateGradesPermission(userId, username))
                {
                    return Forbid();
                }

                if (studentFile == null)
                    return BadRequest(new { error = "Cần cung cấp file: studentFile" });

                if (!IsExcelFile(studentFile))
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });

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
                return StatusCode(500, new { error = "Lỗi hệ thống khi chấm điểm" });
            }
        }

        /// <summary>
        /// Chấm điểm dự án 05
        /// Chi Teacher va Admin duoc phep
        /// </summary>
        [HttpPost("excel/project05")]
        [HttpPost("project05")] // Legacy alias
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject05(
            [FromForm] IFormFile studentFile,
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                if (!HasCreateGradesPermission(userId, username))
                {
                    return Forbid();
                }

                if (studentFile == null)
                    return BadRequest(new { error = "Cần cung cấp file: studentFile" });

                if (!IsExcelFile(studentFile))
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject05Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project05,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading project05");
                return StatusCode(500, new { error = "Lỗi hệ thống khi chấm điểm" });
            }
        }

        /// <summary>
        /// Cham diem du an 06
        /// Chi Teacher va Admin duoc phep
        /// </summary>
        [HttpPost("excel/project06")]
        [HttpPost("project06")] // Legacy alias
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject06(
            [FromForm] IFormFile studentFile,
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                if (!HasCreateGradesPermission(userId, username))
                {
                    return Forbid();
                }

                if (studentFile == null)
                    return BadRequest(new { error = "Can cung cap file: studentFile" });

                if (!IsExcelFile(studentFile))
                    return BadRequest(new { error = "File phai co dinh dang .xlsx hoac .xlsm" });

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject06Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project06,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading project06");
                return StatusCode(500, new { error = "Loi he thong khi cham diem" });
            }
        }

        /// <summary>
        /// Cham diem du an 07
        /// Chi Teacher va Admin duoc phep
        /// </summary>
        [HttpPost("excel/project07")]
        [HttpPost("project07")] // Legacy alias
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject07(
            [FromForm] IFormFile studentFile,
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                if (!HasCreateGradesPermission(userId, username))
                {
                    return Forbid();
                }

                if (studentFile == null)
                    return BadRequest(new { error = "Can cung cap file: studentFile" });

                if (!IsExcelFile(studentFile))
                    return BadRequest(new { error = "File phai co dinh dang .xlsx hoac .xlsm" });

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject07Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project07,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading project07");
                return StatusCode(500, new { error = "Loi he thong khi cham diem" });
            }
        }

        /// <summary>
        /// Cham diem du an 08
        /// Chi Teacher va Admin duoc phep
        /// </summary>
        [HttpPost("excel/project08")]
        [HttpPost("project08")] // Legacy alias
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject08(
            [FromForm] IFormFile studentFile,
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                if (!HasCreateGradesPermission(userId, username))
                {
                    return Forbid();
                }

                if (studentFile == null)
                    return BadRequest(new { error = "Can cung cap file: studentFile" });

                if (!IsExcelFile(studentFile))
                    return BadRequest(new { error = "File phai co dinh dang .xlsx hoac .xlsm" });

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject08Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project08,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading project08");
                return StatusCode(500, new { error = "Loi he thong khi cham diem" });
            }
        }

        /// <summary>
        /// Chấm điểm Project 09
        /// Chỉ Teacher và Admin mới được phép sử dụng
        /// </summary>
        [HttpPost("excel/project09")]
        [HttpPost("project09")] // Legacy alias
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

                // ✅ KIỂM TRA PERMISSION
                if (!HasCreateGradesPermission(userId, username))
                {
                    return Forbid();
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
        /// Chấm điểm Project 10
        /// Chỉ Teacher và Admin mới được phép sử dụng
        /// </summary>
        [HttpPost("excel/project10")]
        [HttpPost("project10")] // Legacy alias
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject10(
            [FromForm] IFormFile studentFile,
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                if (!HasCreateGradesPermission(userId, username))
                {
                    return Forbid();
                }

                if (studentFile == null)
                {
                    return BadRequest(new { error = "Cần cung cấp file: studentFile" });
                }

                if (!IsExcelFile(studentFile))
                {
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });
                }

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject10Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project10,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading project10");
                return StatusCode(500, new { error = "Lỗi hệ thống khi chấm điểm" });
            }
        }

        /// <summary>
        /// Chấm điểm Project 11
        /// Chỉ Teacher và Admin mới được phép sử dụng
        /// </summary>
        [HttpPost("excel/project11")]
        [HttpPost("project11")] // Legacy alias
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject11(
            [FromForm] IFormFile studentFile,
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                if (!HasCreateGradesPermission(userId, username))
                {
                    return Forbid();
                }

                if (studentFile == null)
                {
                    return BadRequest(new { error = "Cần cung cấp file: studentFile" });
                }

                if (!IsExcelFile(studentFile))
                {
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });
                }

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject11Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project11,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading project11");
                return StatusCode(500, new { error = "Lỗi hệ thống khi chấm điểm" });
            }
        }

        /// <summary>
        /// Chấm điểm Project 12
        /// Chỉ Teacher và Admin mới được phép sử dụng
        /// </summary>
        [HttpPost("excel/project12")]
        [HttpPost("project12")] // Legacy alias
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject12(
            [FromForm] IFormFile studentFile,
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                if (!HasCreateGradesPermission(userId, username))
                {
                    return Forbid();
                }

                if (studentFile == null)
                {
                    return BadRequest(new { error = "Cần cung cấp file: studentFile" });
                }

                if (!IsExcelFile(studentFile))
                {
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });
                }

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject12Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project12,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading project12");
                return StatusCode(500, new { error = "Lỗi hệ thống khi chấm điểm" });
            }
        }

        /// <summary>
        /// Chấm điểm Project 13
        /// Chỉ Teacher và Admin mới được phép sử dụng
        /// </summary>
        [HttpPost("excel/project13")]
        [HttpPost("project13")] // Legacy alias
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject13(
            [FromForm] IFormFile studentFile,
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                if (!HasCreateGradesPermission(userId, username))
                {
                    return Forbid();
                }

                if (studentFile == null)
                {
                    return BadRequest(new { error = "Cần cung cấp file: studentFile" });
                }

                if (!IsExcelFile(studentFile))
                {
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });
                }

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject13Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project13,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading project13");
                return StatusCode(500, new { error = "Lỗi hệ thống khi chấm điểm" });
            }
        }

        /// <summary>
        /// Chấm điểm Project 14
        /// Chỉ Teacher và Admin mới được phép sử dụng
        /// </summary>
        [HttpPost("excel/project14")]
        [HttpPost("project14")] // Legacy alias
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject14(
            [FromForm] IFormFile studentFile,
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                if (!HasCreateGradesPermission(userId, username))
                {
                    return Forbid();
                }

                if (studentFile == null)
                {
                    return BadRequest(new { error = "Cần cung cấp file: studentFile" });
                }

                if (!IsExcelFile(studentFile))
                {
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });
                }

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject14Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project14,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading project14");
                return StatusCode(500, new { error = "Lỗi hệ thống khi chấm điểm" });
            }
        }

        /// <summary>
        /// Chấm điểm Project 15
        /// Chỉ Teacher và Admin mới được phép sử dụng
        /// </summary>
        [HttpPost("excel/project15")]
        [HttpPost("project15")] // Legacy alias
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject15(
            [FromForm] IFormFile studentFile,
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                if (!HasCreateGradesPermission(userId, username))
                {
                    return Forbid();
                }

                if (studentFile == null)
                {
                    return BadRequest(new { error = "Cần cung cấp file: studentFile" });
                }

                if (!IsExcelFile(studentFile))
                {
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });
                }

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject15Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project15,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading project15");
                return StatusCode(500, new { error = "Lỗi hệ thống khi chấm điểm" });
            }
        }

        /// <summary>
        /// Chấm điểm Project 16
        /// Chỉ Teacher và Admin mới được phép sử dụng
        /// </summary>
        [HttpPost("excel/project16")]
        [HttpPost("project16")] // Legacy alias
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject16(
            [FromForm] IFormFile studentFile,
            [FromForm] string? classId = null,
            [FromForm] string? assignmentId = null,
            [FromForm] string? studentId = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                if (!HasCreateGradesPermission(userId, username))
                {
                    return Forbid();
                }

                if (studentFile == null)
                {
                    return BadRequest(new { error = "Cần cung cấp file: studentFile" });
                }

                if (!IsExcelFile(studentFile))
                {
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });
                }

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeProject16Async(studentStream);

                await _analyticsService.SaveGradingAttemptAsync(
                    result,
                    GradingApiEndpoints.Project16,
                    classId,
                    assignmentId,
                    studentId,
                    userId);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading project16");
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

        private bool HasCreateGradesPermission(string userId, string username)
        {
            var hasPermission = User.Claims.Any(c =>
                c.Type == "permission" && c.Value == Permissions.CreateGrades);

            if (!hasPermission)
            {
                _logger.LogWarning(
                    "User {Username} (ID: {UserId}) không có quyền {Permission}",
                    username,
                    userId,
                    Permissions.CreateGrades);
            }

            return hasPermission;
        }

        private bool IsExcelFile(IFormFile file)
        {
            var extension = Path.GetExtension(file.FileName).ToLower();
            return extension == ".xlsx" || extension == ".xlsm";
        }
    }
}



