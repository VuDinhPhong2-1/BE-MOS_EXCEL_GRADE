using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Caching.Distributed;
using Microsoft.Extensions.Options;
using MOS.ExcelGrading.API.Helpers;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Security.Claims;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize] // ✅ YÊU CẦU ĐĂNG NHẬP CHO TẤT CẢ ENDPOINT
    public class SchoolController : ControllerBase
    {
        private readonly ISchoolService _schoolService;
        private readonly IDistributedCache _cache;
        private readonly RedisSettings _redisSettings;
        private readonly ILogger<SchoolController> _logger;

        public SchoolController(
            ISchoolService schoolService,
            IDistributedCache cache,
            IOptions<RedisSettings> redisOptions,
            ILogger<SchoolController> logger)
        {
            _schoolService = schoolService;
            _cache = cache;
            _redisSettings = redisOptions.Value;
            _logger = logger;
        }

        /// <summary>
        /// Lấy danh sách schools
        /// Admin: Tất cả schools
        /// Teacher: Tất cả schools (dùng chung)
        /// Student: Chỉ schools mà mình được gán vào
        /// </summary>
        [HttpGet]
        [Authorize(Roles = $"{UserRoles.Admin},{UserRoles.Teacher}")]
        public async Task<IActionResult> GetSchools([FromQuery] bool includeInactive = false)
        {
            try
            {

                // ✅ LẤY THÔNG TIN USER TỪ TOKEN
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
                var userEmail = User.FindFirst(ClaimTypes.Email)?.Value ?? "unknown";

                // ✅ LOG THÔNG TIN USER
                _logger.LogInformation(
                    $"[USER INFO] UserId: {userId}, Username: {username}, Email: {userEmail}, Role: {userRole}, IsAuthenticated: {User.Identity?.IsAuthenticated ?? false}");

                // ✅ KIỂM TRA PERMISSION
                var hasPermission = User.Claims.Any(c =>
                    c.Type == "permission" && c.Value == Permissions.ViewSchools);

                // ✅ LOG PERMISSION CHECK
                _logger.LogInformation(
                    $"[PERMISSION CHECK] User {username} (ID: {userId}) - Permission '{Permissions.ViewSchools}': {hasPermission}");

                if (!hasPermission)
                {
                    _logger.LogWarning(
                        $"[PERMISSION DENIED] User {username} (ID: {userId}) không có quyền {Permissions.ViewSchools}");
                    return Forbid();
                }

                var cacheKey = $"schools:list:v1:user:{userId}:role:{userRole}:inactive:{includeInactive}";
                if (_redisSettings.Enabled)
                {
                    try
                    {
                        _logger.LogInformation("[CACHE LOOKUP] schools key={CacheKey}", cacheKey);
                        var cachedResponse = await _cache.GetJsonAsync<List<SchoolResponse>>(cacheKey);
                        if (cachedResponse != null)
                        {
                            _logger.LogInformation(
                                "[CACHE HIT] schools key={CacheKey}, count={Count}",
                                cacheKey,
                                cachedResponse.Count);
                            return Ok(cachedResponse);
                        }

                        _logger.LogInformation("[CACHE MISS] schools key={CacheKey}", cacheKey);
                    }
                    catch (Exception cacheEx)
                    {
                        _logger.LogWarning(cacheEx, "[CACHE] Không thể đọc cache danh sách trường");
                    }
                }

                // ✅ LẤY DANH SÁCH SCHOOLS
                List<School> schools;

                schools = await _schoolService.GetAllSchoolsAsync(includeInactive);
                _logger.LogInformation(
                    $"[GET SCHOOLS] {userRole} {username} (ID: {userId}) lấy {schools.Count} schools (includeInactive: {includeInactive})");

                // ✅ TẠO RESPONSE
                var response = schools.Select(s => new SchoolResponse
                {
                    Id = s.Id ?? string.Empty,
                    Name = s.Name,
                    Code = s.Code,
                    Address = s.Address,
                    PhoneNumber = s.PhoneNumber,
                    Email = s.Email,
                    Website = s.Website,
                    Description = s.Description,
                    Logo = s.Logo,
                    AttendanceSpreadsheetId = s.AttendanceSpreadsheetId,
                    OwnerId = s.OwnerId,
                    CreatedAt = s.CreatedAt,
                    IsActive = s.IsActive
                }).ToList();

                // ✅ LOG THÀNH CÔNG
                _logger.LogInformation(
                    $"[GET SCHOOLS SUCCESS] User {username} (ID: {userId}) - Trả về {response.Count} schools");

                if (_redisSettings.Enabled)
                {
                    try
                    {
                        var ttl = ResolveTtl(_redisSettings.SchoolsTtlSeconds);
                        await _cache.SetJsonAsync(
                            cacheKey,
                            response,
                            ttl);
                        _logger.LogInformation(
                            "[CACHE SET] schools key={CacheKey}, count={Count}, ttlSeconds={TtlSeconds}",
                            cacheKey,
                            response.Count,
                            ttl.TotalSeconds);
                    }
                    catch (Exception cacheEx)
                    {
                        _logger.LogWarning(cacheEx, "[CACHE] Không thể ghi cache danh sách trường");
                    }
                }

                return Ok(response);
            }
            catch (Exception ex)
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                _logger.LogError(ex,
                    $"[GET SCHOOLS ERROR] User {username} (ID: {userId}) - Lỗi: {ex.Message}");

                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }
        /// <summary>
        /// Lấy school theo ID
        /// Admin: Xem tất cả
        /// Teacher/Student: Xem school dùng chung
        /// </summary>
        [HttpGet("{id}")]
        [Authorize(Roles = $"{UserRoles.Admin},{UserRoles.Teacher}")]
        public async Task<IActionResult> GetSchoolById(string id)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;

                // ✅ KIỂM TRA PERMISSION
                var hasPermission = User.Claims.Any(c =>
                    c.Type == "permission" && c.Value == Permissions.ViewSchools);

                if (!hasPermission)
                {
                    _logger.LogWarning($"User {username} (ID: {userId}) không có quyền {Permissions.ViewSchools}");
                    return Forbid();
                }

                var school = await _schoolService.GetSchoolByIdAsync(id);

                if (school == null)
                {
                    _logger.LogWarning($"[GET SCHOOL] User {username} tìm school ID {id} không tồn tại");
                    return NotFound(new { message = "Không tìm thấy trường" });
                }

                _logger.LogInformation($"[GET SCHOOL] User {username} xem school {school.Name} (ID: {id})");

                var response = new SchoolResponse
                {
                    Id = school.Id ?? string.Empty,
                    Name = school.Name,
                    Code = school.Code,
                    Address = school.Address,
                    PhoneNumber = school.PhoneNumber,
                    Email = school.Email,
                    Website = school.Website,
                    Description = school.Description,
                    Logo = school.Logo,
                    AttendanceSpreadsheetId = school.AttendanceSpreadsheetId,
                    OwnerId = school.OwnerId,
                    CreatedAt = school.CreatedAt,
                    IsActive = school.IsActive
                };

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi lấy thông tin school");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }

        /// <summary>
        /// Tạo school mới
        /// Chỉ Teacher và Admin mới được tạo
        /// Teacher tự động trở thành owner
        /// </summary>
        [HttpPost]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> CreateSchool([FromBody] CreateSchoolRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
                var isAdmin = string.Equals(userRole, UserRoles.Admin, StringComparison.Ordinal);

                // ✅ KIỂM TRA PERMISSION
                var hasPermission = User.Claims.Any(c =>
                    c.Type == "permission" && c.Value == Permissions.CreateSchools);

                if (!hasPermission)
                {
                    _logger.LogWarning($"User {username} (ID: {userId}) không có quyền {Permissions.CreateSchools}");
                    return Forbid();
                }

                var normalizedCode = NormalizeSchoolCode(request.Code);
                if (string.IsNullOrWhiteSpace(normalizedCode))
                {
                    return BadRequest(new { message = "Mã trường là bắt buộc" });
                }

                // Kiểm tra mã trường đã tồn tại chưa
                if (await _schoolService.SchoolExistsAsync(normalizedCode))
                {
                    _logger.LogWarning($"[CREATE SCHOOL] User {username} tạo school với mã {normalizedCode} đã tồn tại");
                    return BadRequest(new { message = "Mã trường đã tồn tại" });
                }

                var normalizedSpreadsheetId = NormalizeOptional(request.AttendanceSpreadsheetId);
                if (!isAdmin && !string.IsNullOrWhiteSpace(normalizedSpreadsheetId))
                {
                    _logger.LogWarning(
                        "[CREATE SCHOOL] User {Username} (ID: {UserId}, Role: {Role}) cố gắng set AttendanceSpreadsheetId nhưng không đủ quyền",
                        username,
                        userId,
                        userRole);
                    return StatusCode(StatusCodes.Status403Forbidden, new
                    {
                        message = "Chỉ Admin mới được cấu hình Spreadsheet ID Google Sheet."
                    });
                }

                var school = new School
                {
                    Name = request.Name,
                    Code = normalizedCode,
                    Address = request.Address,
                    PhoneNumber = request.PhoneNumber,
                    Email = request.Email,
                    Website = request.Website,
                    Description = request.Description,
                    Logo = request.Logo,
                    AttendanceSpreadsheetId = isAdmin ? normalizedSpreadsheetId : null
                };

                var createdSchool = await _schoolService.CreateSchoolAsync(school, userId);

                _logger.LogInformation(
                    $"[CREATE SCHOOL] User {username} (ID: {userId}, Role: {userRole}) " +
                    $"tạo school {createdSchool.Name} (Code: {createdSchool.Code}, ID: {createdSchool.Id})");

                return CreatedAtAction(
                    nameof(GetSchoolById),
                    new { id = createdSchool.Id },
                    new SchoolResponse
                    {
                        Id = createdSchool.Id ?? string.Empty,
                        Name = createdSchool.Name,
                        Code = createdSchool.Code,
                        Address = createdSchool.Address,
                        PhoneNumber = createdSchool.PhoneNumber,
                        Email = createdSchool.Email,
                        Website = createdSchool.Website,
                        Description = createdSchool.Description,
                        Logo = createdSchool.Logo,
                        AttendanceSpreadsheetId = createdSchool.AttendanceSpreadsheetId,
                        OwnerId = createdSchool.OwnerId,
                        CreatedAt = createdSchool.CreatedAt,
                        IsActive = createdSchool.IsActive
                    }
                );
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi tạo school");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }

        /// <summary>
        /// Cập nhật school
        /// Teacher và Admin được phép
        /// </summary>
        [HttpPut("{id}")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> UpdateSchool(string id, [FromBody] UpdateSchoolRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
                var isAdmin = string.Equals(userRole, UserRoles.Admin, StringComparison.Ordinal);

                // ✅ KIỂM TRA PERMISSION
                var hasPermission = User.Claims.Any(c =>
                    c.Type == "permission" && c.Value == Permissions.EditSchools);

                if (!hasPermission)
                {
                    _logger.LogWarning($"User {username} (ID: {userId}) không có quyền {Permissions.EditSchools}");
                    return Forbid();
                }

                var existingSchool = await _schoolService.GetSchoolByIdAsync(id);
                if (existingSchool == null)
                {
                    _logger.LogWarning($"[UPDATE SCHOOL] User {username} cập nhật school ID {id} không tồn tại");
                    return NotFound(new { message = "Không tìm thấy trường" });
                }

                // Cập nhật các field
                if (!string.IsNullOrEmpty(request.Name))
                    existingSchool.Name = request.Name;

                if (!string.IsNullOrWhiteSpace(request.Code))
                {
                    var normalizedCode = NormalizeSchoolCode(request.Code);
                    if (string.IsNullOrWhiteSpace(normalizedCode))
                    {
                        return BadRequest(new { message = "Mã trường không hợp lệ" });
                    }

                    if (!string.Equals(existingSchool.Code, normalizedCode, StringComparison.OrdinalIgnoreCase)
                        && await _schoolService.SchoolExistsAsync(normalizedCode))
                    {
                        return BadRequest(new { message = "Mã trường đã tồn tại" });
                    }

                    existingSchool.Code = normalizedCode;
                }

                existingSchool.Address = request.Address ?? existingSchool.Address;
                existingSchool.PhoneNumber = request.PhoneNumber ?? existingSchool.PhoneNumber;
                existingSchool.Email = request.Email ?? existingSchool.Email;
                existingSchool.Website = request.Website ?? existingSchool.Website;
                existingSchool.Description = request.Description ?? existingSchool.Description;
                existingSchool.Logo = request.Logo ?? existingSchool.Logo;
                if (request.AttendanceSpreadsheetId != null)
                {
                    var normalizedRequestedSpreadsheetId = NormalizeOptional(request.AttendanceSpreadsheetId);
                    var normalizedCurrentSpreadsheetId = NormalizeOptional(existingSchool.AttendanceSpreadsheetId);
                    var isSpreadsheetChanged = !string.Equals(
                        normalizedRequestedSpreadsheetId,
                        normalizedCurrentSpreadsheetId,
                        StringComparison.Ordinal);

                    if (isSpreadsheetChanged && !isAdmin)
                    {
                        _logger.LogWarning(
                            "[UPDATE SCHOOL] User {Username} (ID: {UserId}, Role: {Role}) cố gắng đổi AttendanceSpreadsheetId của school {SchoolId} nhưng không đủ quyền",
                            username,
                            userId,
                            userRole,
                            id);
                        return StatusCode(StatusCodes.Status403Forbidden, new
                        {
                            message = "Chỉ Admin mới được thay đổi Spreadsheet ID Google Sheet."
                        });
                    }

                    if (isAdmin)
                    {
                        existingSchool.AttendanceSpreadsheetId = normalizedRequestedSpreadsheetId;
                    }
                }

                if (request.IsActive.HasValue)
                    existingSchool.IsActive = request.IsActive.Value;

                var updatedSchool = await _schoolService.UpdateSchoolAsync(id, existingSchool, userId);

                if (updatedSchool == null)
                {
                    _logger.LogError($"[UPDATE SCHOOL] Không thể cập nhật school {id}");
                    return NotFound(new { message = "Không thể cập nhật trường" });
                }

                _logger.LogInformation(
                    $"[UPDATE SCHOOL] User {username} (ID: {userId}, Role: {userRole}) " +
                    $"cập nhật school {updatedSchool.Name} (ID: {id})");

                return Ok(new SchoolResponse
                {
                    Id = updatedSchool.Id ?? string.Empty,
                    Name = updatedSchool.Name,
                    Code = updatedSchool.Code,
                    Address = updatedSchool.Address,
                    PhoneNumber = updatedSchool.PhoneNumber,
                    Email = updatedSchool.Email,
                    Website = updatedSchool.Website,
                    Description = updatedSchool.Description,
                    Logo = updatedSchool.Logo,
                    AttendanceSpreadsheetId = updatedSchool.AttendanceSpreadsheetId,
                    OwnerId = updatedSchool.OwnerId,
                    CreatedAt = updatedSchool.CreatedAt,
                    IsActive = updatedSchool.IsActive
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi cập nhật school");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }

        /// <summary>
        /// Xóa school
        /// Chỉ Admin mới được phép
        /// </summary>
        [HttpDelete("{id}")]
        [Authorize(Roles = $"{UserRoles.Admin}")]
        public async Task<IActionResult> DeleteSchool(string id)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;

                // ✅ KIỂM TRA PERMISSION
                var hasPermission = User.Claims.Any(c =>
                    c.Type == "permission" && c.Value == Permissions.DeleteSchools);

                if (!hasPermission)
                {
                    _logger.LogWarning($"User {username} (ID: {userId}) không có quyền {Permissions.DeleteSchools}");
                    return Forbid();
                }

                var school = await _schoolService.GetSchoolByIdAsync(id);
                if (school == null)
                {
                    _logger.LogWarning($"[DELETE SCHOOL] User {username} xóa school ID {id} không tồn tại");
                    return NotFound(new { message = "Không tìm thấy trường" });
                }

                var result = await _schoolService.DeleteSchoolAsync(id);

                if (!result)
                {
                    _logger.LogError($"[DELETE SCHOOL] Không thể xóa school {id}");
                    return BadRequest(new { message = "Không thể xóa trường" });
                }

                _logger.LogInformation(
                    $"[DELETE SCHOOL] User {username} (ID: {userId}, Role: {userRole}) " +
                    $"xóa school {school.Name} (ID: {id})");

                return Ok(new { message = "Đã xóa trường thành công" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi xóa school");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }

        private static string NormalizeSchoolCode(string? code)
        {
            if (string.IsNullOrWhiteSpace(code))
                return string.Empty;

            var normalized = new string(code
                .Trim()
                .ToUpperInvariant()
                .Where(c => char.IsLetterOrDigit(c) || c == '-' || c == '_')
                .ToArray());

            return normalized;
        }

        private static string? NormalizeOptional(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return null;

            return value.Trim();
        }

        private static TimeSpan ResolveTtl(int configuredSeconds, int fallbackSeconds = 60)
        {
            var safeSeconds = configuredSeconds > 0 ? configuredSeconds : fallbackSeconds;
            return TimeSpan.FromSeconds(safeSeconds);
        }
    }
}
