using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
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
    public class AuthController : ControllerBase
    {
        private readonly IUserService _userService;
        private readonly IDistributedCache _cache;
        private readonly RedisSettings _redisSettings;
        private readonly ILogger<AuthController> _logger;

        public AuthController(
            IUserService userService,
            IDistributedCache cache,
            IOptions<RedisSettings> redisOptions,
            ILogger<AuthController> logger)
        {
            _userService = userService;
            _cache = cache;
            _redisSettings = redisOptions.Value;
            _logger = logger;
        }

        /// <summary>
        /// Đăng ký tài khoản mới
        /// </summary>
        [HttpPost("register")]
        public async Task<IActionResult> Register([FromBody] RegisterRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                // ========== CHẶN ĐĂNG KÝ ADMIN ==========
                if (!string.IsNullOrEmpty(request.Role) && request.Role == UserRoles.Admin)
                {
                    return BadRequest(new { message = "Không thể đăng ký tài khoản Admin qua API" });
                }

                var user = await _userService.RegisterAsync(
                    request.Email,
                    request.Username,
                    request.Password,
                    request.Role,
                    request.FullName
                );

                if (user == null)
                    return BadRequest(new { message = "Tên đăng nhập hoặc thư điện tử đã tồn tại" });

                return Ok(new
                {
                    message = "Đăng ký thành công",
                    userId = user.Id,
                    username = user.Username,
                    email = user.Email,
                    role = user.Role,
                    permissions = user.Permissions
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi đăng ký user");
                return StatusCode(500, new { message = "Đã xảy ra lỗi khi đăng ký" });
            }
        }

        /// <summary>
        /// Đăng nhập - Trả về Access Token và Refresh Token
        /// </summary>
        [HttpPost("login")]
        public async Task<IActionResult> Login([FromBody] LoginRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                var authResponse = await _userService.LoginAsync(request.Username, request.Password);

                if (authResponse == null)
                    return Unauthorized(new { message = "Tên đăng nhập hoặc mật khẩu không đúng hoặc tài khoản đã bị vô hiệu hóa" });

                return Ok(authResponse);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi đăng nhập");
                return StatusCode(500, new { message = "Đã xảy ra lỗi khi đăng nhập" });
            }
        }

        /// <summary>
        /// Đăng nhập bằng Google ID Token
        /// </summary>
        [HttpPost("google-login")]
        public async Task<IActionResult> GoogleLogin([FromBody] GoogleLoginRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                var authResponse = await _userService.LoginWithGoogleAsync(request.IdToken);

                if (authResponse == null)
                    return Unauthorized(new { message = "Mã thông báo Google không hợp lệ hoặc tài khoản đã bị vô hiệu hóa" });

                return Ok(authResponse);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi đăng nhập Google");
                return StatusCode(500, new { message = "Đã xảy ra lỗi khi đăng nhập Google" });
            }
        }

        /// <summary>
        /// Làm mới Access Token bằng Refresh Token
        /// </summary>
        [HttpPost("refresh-token")]
        public async Task<IActionResult> RefreshToken([FromBody] RefreshTokenRequest request)
        {
            try
            {
                if (string.IsNullOrEmpty(request.RefreshToken))
                    return BadRequest(new { message = "Mã làm mới phiên là bắt buộc" });

                var response = await _userService.RefreshTokenAsync(request.RefreshToken);

                if (response == null)
                {
                    return Unauthorized(new { message = "Mã làm mới phiên không hợp lệ hoặc đã hết hạn" });
                }

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi refresh token");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }

        /// <summary>
        /// Đăng xuất - Thu hồi Refresh Token
        /// </summary>
        [HttpPost("logout")]
        [Authorize]
        public async Task<IActionResult> Logout()
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;

                if (string.IsNullOrEmpty(userId))
                    return Unauthorized(new { message = "Mã xác thực không hợp lệ" });

                var result = await _userService.RevokeRefreshTokenAsync(userId);

                if (!result)
                    return BadRequest(new { message = "Không thể đăng xuất" });

                return Ok(new { message = "Đăng xuất thành công" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi đăng xuất");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }

        /// <summary>
        /// Lấy thông tin user hiện tại (cần token)
        /// </summary>
        [HttpGet("me")]
        [Authorize]
        public async Task<IActionResult> GetCurrentUser()
        {
            try
            {
                var username = User.Identity?.Name;
                if (string.IsNullOrEmpty(username))
                    return Unauthorized(new { message = "Mã xác thực không hợp lệ" });

                var user = await _userService.GetUserByUsernameAsync(username);

                if (user == null)
                    return NotFound(new { message = "Không tìm thấy người dùng" });

                return Ok(new
                {
                    userId = user.Id,
                    username = user.Username,
                    email = user.Email,
                    fullName = user.FullName,
                    phoneNumber = user.PhoneNumber,
                    role = user.Role,
                    permissions = user.Permissions,
                    avatar = user.Avatar,
                    createdAt = user.CreatedAt,
                    lastLogin = user.LastLogin,
                    isActive = user.IsActive
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi lấy thông tin user");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }

        /// <summary>
        /// Lấy danh sách giáo viên để gán quyền/bàn giao
        /// </summary>
        [HttpGet("teachers")]
        [Authorize(Roles = $"{UserRoles.Admin},{UserRoles.Teacher}")]
        public async Task<IActionResult> GetTeachers([FromQuery] bool includeInactive = false)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var username = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                var hasPermission = User.Claims.Any(c =>
                    c.Type == "permission" && c.Value == Permissions.ViewUsers);

                if (!hasPermission)
                {
                    _logger.LogWarning($"User {username} (ID: {userId}) không có quyền {Permissions.ViewUsers}");
                    return Forbid();
                }

                var cacheKey = $"teachers:list:v1:user:{userId}:inactive:{includeInactive}";
                if (_redisSettings.Enabled)
                {
                    try
                    {
                        var cachedResponse = await _cache.GetJsonAsync<List<TeacherListItemResponse>>(cacheKey);
                        if (cachedResponse != null)
                        {
                            return Ok(cachedResponse);
                        }
                    }
                    catch (Exception cacheEx)
                    {
                        _logger.LogWarning(cacheEx, "[CACHE] Không thể đọc cache danh sách giáo viên");
                    }
                }

                var teachers = await _userService.GetTeachersAsync(includeInactive);
                var response = teachers.Select(t => new TeacherListItemResponse
                {
                    UserId = t.Id ?? string.Empty,
                    Username = t.Username ?? string.Empty,
                    FullName = t.FullName ?? string.Empty,
                    Email = t.Email ?? string.Empty,
                    Role = t.Role ?? string.Empty,
                    Permissions = t.Permissions ?? new List<string>(),
                    IsActive = t.IsActive
                }).ToList();

                if (_redisSettings.Enabled)
                {
                    try
                    {
                        await _cache.SetJsonAsync(
                            cacheKey,
                            response,
                            ResolveTtl(_redisSettings.TeachersTtlSeconds));
                    }
                    catch (Exception cacheEx)
                    {
                        _logger.LogWarning(cacheEx, "[CACHE] Không thể ghi cache danh sách giáo viên");
                    }
                }

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi lấy danh sách giáo viên");
                return StatusCode(500, new { message = "Đã xảy ra lỗi khi lấy danh sách giáo viên" });
            }
        }

        /// <summary>
        /// Lấy danh sách permission để Admin phân quyền giáo viên
        /// </summary>
        [HttpGet("permissions")]
        [Authorize(Roles = $"{UserRoles.Admin}")]
        public IActionResult GetPermissionCatalog()
        {
            var permissionCatalog = Permissions.GetRolePermissions()
                .SelectMany(item => item.Value)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(item => item, StringComparer.Ordinal)
                .ToList();

            var teacherDefaults = Permissions.GetRolePermissions().TryGetValue(UserRoles.Teacher, out var defaults)
                ? defaults
                : new List<string>();

            return Ok(new
            {
                permissions = permissionCatalog,
                teacherDefaultPermissions = teacherDefaults
            });
        }

        /// <summary>
        /// Admin cập nhật permission cho giáo viên
        /// </summary>
        [HttpPut("teachers/{teacherId}/permissions")]
        [Authorize(Roles = $"{UserRoles.Admin}")]
        public async Task<IActionResult> UpdateTeacherPermissions(string teacherId, [FromBody] UpdateUserPermissionsRequest request)
        {
            try
            {
                var adminUserId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "unknown";
                var adminUsername = User.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";

                var hasPermission = User.Claims.Any(c =>
                    c.Type == "permission" && c.Value == Permissions.EditUsers);

                if (!hasPermission)
                {
                    _logger.LogWarning($"User {adminUsername} (ID: {adminUserId}) không có quyền {Permissions.EditUsers}");
                    return Forbid();
                }

                if (string.IsNullOrWhiteSpace(teacherId))
                    return BadRequest(new { message = "TeacherId là bắt buộc" });

                if (request == null)
                    return BadRequest(new { message = "Dữ liệu phân quyền không hợp lệ" });

                var updatedTeacher = await _userService.UpdateTeacherPermissionsAsync(teacherId, request.Permissions);
                if (updatedTeacher == null)
                    return NotFound(new { message = "Không tìm thấy giáo viên để cập nhật quyền" });

                _logger.LogInformation(
                    "[UPDATE TEACHER PERMISSIONS] Admin {AdminUsername} (ID: {AdminUserId}) cập nhật quyền cho teacher {TeacherId}",
                    adminUsername,
                    adminUserId,
                    teacherId);

                return Ok(new
                {
                    userId = updatedTeacher.Id ?? string.Empty,
                    username = updatedTeacher.Username,
                    fullName = updatedTeacher.FullName,
                    email = updatedTeacher.Email,
                    role = updatedTeacher.Role,
                    permissions = updatedTeacher.Permissions,
                    isActive = updatedTeacher.IsActive
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi cập nhật permission cho giáo viên");
                return StatusCode(500, new { message = "Đã xảy ra lỗi khi cập nhật phân quyền" });
            }
        }

        /// <summary>
        /// Cập nhật thông tin tài khoản hiện tại
        /// </summary>
        [HttpPut("me")]
        [Authorize]
        public async Task<IActionResult> UpdateCurrentUserProfile([FromBody] UpdateProfileRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrWhiteSpace(userId))
                    return Unauthorized(new { message = "Mã xác thực không hợp lệ" });

                var updatedUser = await _userService.UpdateProfileAsync(userId, request);
                if (updatedUser == null)
                    return NotFound(new { message = "Không tìm thấy người dùng hoặc tài khoản đã bị vô hiệu hóa" });

                return Ok(new
                {
                    userId = updatedUser.Id,
                    username = updatedUser.Username,
                    email = updatedUser.Email,
                    fullName = updatedUser.FullName,
                    phoneNumber = updatedUser.PhoneNumber,
                    avatar = updatedUser.Avatar,
                    role = updatedUser.Role,
                    permissions = updatedUser.Permissions,
                    isActive = updatedUser.IsActive
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi cập nhật hồ sơ người dùng");
                return StatusCode(500, new { message = "Đã xảy ra lỗi khi cập nhật thông tin tài khoản" });
            }
        }

        private static TimeSpan ResolveTtl(int configuredSeconds, int fallbackSeconds = 60)
        {
            var safeSeconds = configuredSeconds > 0 ? configuredSeconds : fallbackSeconds;
            return TimeSpan.FromSeconds(safeSeconds);
        }

        private sealed class TeacherListItemResponse
        {
            public string UserId { get; init; } = string.Empty;
            public string Username { get; init; } = string.Empty;
            public string FullName { get; init; } = string.Empty;
            public string Email { get; init; } = string.Empty;
            public string Role { get; init; } = string.Empty;
            public List<string> Permissions { get; init; } = new();
            public bool IsActive { get; init; }
        }
    }
}


