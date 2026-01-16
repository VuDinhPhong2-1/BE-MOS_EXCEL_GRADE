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
    [Authorize] // ✅ YÊU CẦU ĐĂNG NHẬP CHO TẤT CẢ ENDPOINT
    public class SchoolController : ControllerBase
    {
        private readonly ISchoolService _schoolService;
        private readonly ILogger<SchoolController> _logger;

        public SchoolController(ISchoolService schoolService, ILogger<SchoolController> logger)
        {
            _schoolService = schoolService;
            _logger = logger;
        }

        /// <summary>
        /// Lấy danh sách schools
        /// Admin: Tất cả schools
        /// Teacher: Chỉ schools mà mình tạo ra
        /// Student: Chỉ schools mà mình được gán vào
        /// </summary>
        [HttpGet]
        [Authorize(Roles = $"{UserRoles.Admin},{UserRoles.Teacher}")]
        /// <summary>
        /// Lấy danh sách schools
        /// Admin: Tất cả schools
        /// Teacher: Chỉ schools mà mình tạo ra
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

                // ✅ LẤY DANH SÁCH SCHOOLS
                List<School> schools;

                if (userRole == UserRoles.Admin)
                {
                    schools = await _schoolService.GetAllSchoolsAsync(includeInactive);
                    _logger.LogInformation(
                        $"[GET SCHOOLS] Admin {username} (ID: {userId}) lấy tất cả {schools.Count} schools (includeInactive: {includeInactive})");
                }
                else
                {
                    schools = await _schoolService.GetSchoolsByOwnerIdAsync(userId, includeInactive);
                    _logger.LogInformation(
                        $"[GET SCHOOLS] {userRole} {username} (ID: {userId}) lấy {schools.Count} schools của mình (includeInactive: {includeInactive})");
                }

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
                    OwnerId = s.OwnerId,
                    CreatedAt = s.CreatedAt,
                    IsActive = s.IsActive
                }).ToList();

                // ✅ LOG THÀNH CÔNG
                _logger.LogInformation(
                    $"[GET SCHOOLS SUCCESS] User {username} (ID: {userId}) - Trả về {response.Count} schools");

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
        /// Teacher/Student: Chỉ xem school mà mình là owner hoặc được gán vào
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

                // ✅ KIỂM TRA QUYỀN TRUY CẬP
                if (userRole != UserRoles.Admin && school.OwnerId != userId)
                {
                    _logger.LogWarning($"[GET SCHOOL] User {username} (ID: {userId}) không có quyền xem school {id}");
                    return Forbid();
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

                // ✅ KIỂM TRA PERMISSION
                var hasPermission = User.Claims.Any(c =>
                    c.Type == "permission" && c.Value == Permissions.CreateSchools);

                if (!hasPermission)
                {
                    _logger.LogWarning($"User {username} (ID: {userId}) không có quyền {Permissions.CreateSchools}");
                    return Forbid();
                }

                // Kiểm tra mã trường đã tồn tại chưa
                if (await _schoolService.SchoolExistsAsync(request.Code))
                {
                    _logger.LogWarning($"[CREATE SCHOOL] User {username} tạo school với mã {request.Code} đã tồn tại");
                    return BadRequest(new { message = "Mã trường đã tồn tại" });
                }

                var school = new School
                {
                    Name = request.Name,
                    Code = request.Code,
                    Address = request.Address,
                    PhoneNumber = request.PhoneNumber,
                    Email = request.Email,
                    Website = request.Website,
                    Description = request.Description,
                    Logo = request.Logo
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
        /// Chỉ owner hoặc Admin mới được phép
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

                // ✅ KIỂM TRA QUYỀN: Chỉ owner hoặc Admin mới được sửa
                if (userRole != UserRoles.Admin && existingSchool.OwnerId != userId)
                {
                    _logger.LogWarning($"[UPDATE SCHOOL] User {username} (ID: {userId}) không có quyền sửa school {id}");
                    return Forbid();
                }

                // Cập nhật các field
                if (!string.IsNullOrEmpty(request.Name))
                    existingSchool.Name = request.Name;

                if (!string.IsNullOrEmpty(request.Code))
                    existingSchool.Code = request.Code;

                existingSchool.Address = request.Address ?? existingSchool.Address;
                existingSchool.PhoneNumber = request.PhoneNumber ?? existingSchool.PhoneNumber;
                existingSchool.Email = request.Email ?? existingSchool.Email;
                existingSchool.Website = request.Website ?? existingSchool.Website;
                existingSchool.Description = request.Description ?? existingSchool.Description;
                existingSchool.Logo = request.Logo ?? existingSchool.Logo;

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
        /// Chỉ owner hoặc Admin mới được phép
        /// </summary>
        [HttpDelete("{id}")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
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

                // ✅ KIỂM TRA QUYỀN: Chỉ owner hoặc Admin mới được xóa
                if (userRole != UserRoles.Admin && school.OwnerId != userId)
                {
                    _logger.LogWarning($"[DELETE SCHOOL] User {username} (ID: {userId}) không có quyền xóa school {id}");
                    return Forbid();
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
    }
}
