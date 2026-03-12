using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
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
        private readonly ILogger<AuthController> _logger;

        public AuthController(IUserService userService, ILogger<AuthController> logger)
        {
            _userService = userService;
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

                var teachers = await _userService.GetTeachersAsync(includeInactive);
                var response = teachers.Select(t => new
                {
                    userId = t.Id ?? string.Empty,
                    username = t.Username,
                    fullName = t.FullName,
                    email = t.Email,
                    isActive = t.IsActive
                });

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi lấy danh sách giáo viên");
                return StatusCode(500, new { message = "Đã xảy ra lỗi khi lấy danh sách giáo viên" });
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
    }
}


