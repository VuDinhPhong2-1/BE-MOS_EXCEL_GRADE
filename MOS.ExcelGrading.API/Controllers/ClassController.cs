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
    [Authorize]
    public class ClassController : ControllerBase
    {
        private readonly IClassService _classService;
        private readonly ISchoolService _schoolService;
        private readonly ILogger<ClassController> _logger;

        public ClassController(
            IClassService classService,
            ISchoolService schoolService,
            ILogger<ClassController> logger)
        {
            _classService = classService;
            _schoolService = schoolService;
            _logger = logger;
        }

        /// <summary>
        /// Lấy danh sách classes
        /// Admin: Tất cả classes
        /// Teacher: Chỉ classes mà mình tạo ra
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> GetClasses([FromQuery] bool includeInactive = false)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;

                List<Class> classes;

                if (userRole == UserRoles.Admin)
                {
                    classes = await _classService.GetAllClassesAsync(includeInactive);
                }
                else
                {
                    classes = await _classService.GetClassesByOwnerIdAsync(userId, includeInactive);
                }

                var response = classes.Select(c => new ClassResponse
                {
                    Id = c.Id ?? string.Empty,
                    Name = c.Name,
                    SchoolId = c.SchoolId,
                    OwnerId = c.OwnerId,
                    Description = c.Description,
                    MaxStudents = c.MaxStudents,
                    CurrentStudents = c.CurrentStudents,
                    AcademicYear = c.AcademicYear,
                    Grade = c.Grade,
                    StudentIds = c.StudentIds,
                    CreatedAt = c.CreatedAt,
                    IsActive = c.IsActive
                }).ToList();

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi lấy danh sách classes");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }

        /// <summary>
        /// Lấy danh sách classes theo SchoolId
        /// </summary>
        [HttpGet("school/{schoolId}")]
        public async Task<IActionResult> GetClassesBySchoolId(string schoolId, [FromQuery] bool includeInactive = false)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;

                _logger.LogInformation($"📤 GetClassesBySchoolId called: schoolId={schoolId}, userId={userId}, role={userRole}");

                // ✅ KIỂM TRA SCHOOLID HỢP LỆ
                if (string.IsNullOrEmpty(schoolId) || schoolId.Length != 24)
                {
                    _logger.LogWarning($"❌ Invalid schoolId: {schoolId}");
                    return BadRequest(new { message = "SchoolId không hợp lệ" });
                }

                // ✅ KIỂM TRA SCHOOL TỒN TẠI
                School? school = null;
                try
                {
                    school = await _schoolService.GetSchoolByIdAsync(schoolId);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"❌ Error getting school: {schoolId}");
                    return StatusCode(500, new { message = "Lỗi khi lấy thông tin trường", error = ex.Message });
                }

                if (school == null)
                {
                    _logger.LogWarning($"❌ School not found: {schoolId}");
                    return NotFound(new { message = "Không tìm thấy trường" });
                }

                // ✅ KIỂM TRA QUYỀN TRUY CẬP
                if (userRole != UserRoles.Admin && school.OwnerId != userId)
                {
                    _logger.LogWarning($"❌ Access denied: userId={userId}, schoolOwnerId={school.OwnerId}");
                    return Forbid();
                }

                // ✅ LẤY DANH SÁCH LỚP
                List<Class> classes;
                try
                {
                    classes = await _classService.GetClassesBySchoolIdAsync(schoolId, includeInactive);
                    _logger.LogInformation($"✅ Found {classes.Count} classes for school {schoolId}");
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"❌ Error getting classes for school: {schoolId}");
                    return StatusCode(500, new { message = "Lỗi khi lấy danh sách lớp", error = ex.Message });
                }

                var response = classes.Select(c => new ClassResponse
                {
                    Id = c.Id ?? string.Empty,
                    Name = c.Name,
                    SchoolId = c.SchoolId,
                    OwnerId = c.OwnerId,
                    Description = c.Description,
                    MaxStudents = c.MaxStudents,
                    CurrentStudents = c.CurrentStudents,
                    AcademicYear = c.AcademicYear,
                    Grade = c.Grade,
                    StudentIds = c.StudentIds,
                    CreatedAt = c.CreatedAt,
                    IsActive = c.IsActive
                }).ToList();

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"❌ Unexpected error in GetClassesBySchoolId: {schoolId}");
                return StatusCode(500, new { message = "Đã xảy ra lỗi không mong muốn", error = ex.Message });
            }
        }

        /// <summary>
        /// Lấy class theo ID
        /// </summary>
        [HttpGet("{id}")]
        public async Task<IActionResult> GetClassById(string id)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;

                var classEntity = await _classService.GetClassByIdAsync(id);

                if (classEntity == null)
                    return NotFound(new { message = "Không tìm thấy lớp" });

                // Kiểm tra quyền truy cập
                if (userRole != UserRoles.Admin && classEntity.OwnerId != userId)
                {
                    return Forbid();
                }

                var response = new ClassResponse
                {
                    Id = classEntity.Id ?? string.Empty,
                    Name = classEntity.Name,
                    SchoolId = classEntity.SchoolId,
                    OwnerId = classEntity.OwnerId,
                    Description = classEntity.Description,
                    MaxStudents = classEntity.MaxStudents,
                    CurrentStudents = classEntity.CurrentStudents,
                    AcademicYear = classEntity.AcademicYear,
                    Grade = classEntity.Grade,
                    StudentIds = classEntity.StudentIds,
                    CreatedAt = classEntity.CreatedAt,
                    IsActive = classEntity.IsActive
                };

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi lấy thông tin class");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }

        /// <summary>
        /// Tạo class mới
        /// Teacher chỉ tạo được class trong school mà mình là owner
        /// </summary>
        [HttpPost]
        public async Task<IActionResult> CreateClass([FromBody] CreateClassRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;

                // Kiểm tra school tồn tại
                var school = await _schoolService.GetSchoolByIdAsync(request.SchoolId);
                if (school == null)
                    return NotFound(new { message = "Không tìm thấy trường" });

                // Kiểm tra quyền: Teacher chỉ tạo class trong school mà mình là owner
                if (userRole != UserRoles.Admin && school.OwnerId != userId)
                {
                    return Forbid();
                }

                // Kiểm tra mã lớp đã tồn tại trong school chưa
                if (await _classService.ClassExistsAsync( request.Name))
                    return BadRequest(new { message = "Tên lớp đã tồn tại trong trường này" });

                var classEntity = new Class
                {
                    Name = request.Name,
                    SchoolId = request.SchoolId,
                    Description = request.Description,
                    MaxStudents = request.MaxStudents,
                    AcademicYear = request.AcademicYear,
                    Grade = request.Grade,
                };

                var createdClass = await _classService.CreateClassAsync(classEntity, userId);

                return CreatedAtAction(
                    nameof(GetClassById),
                    new { id = createdClass.Id },
                    new ClassResponse
                    {
                        Id = createdClass.Id ?? string.Empty,
                        Name = createdClass.Name,
                        SchoolId = createdClass.SchoolId,
                        OwnerId = createdClass.OwnerId,
                        Description = createdClass.Description,
                        MaxStudents = createdClass.MaxStudents,
                        CurrentStudents = createdClass.CurrentStudents,
                        AcademicYear = createdClass.AcademicYear,
                        Grade = createdClass.Grade,
                        CreatedAt = createdClass.CreatedAt,
                        IsActive = createdClass.IsActive
                    }
                );
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi tạo class");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }

        /// <summary>
        /// Cập nhật class
        /// Chỉ owner hoặc Admin mới được phép
        /// </summary>
        [HttpPut("{id}")]
        public async Task<IActionResult> UpdateClass(string id, [FromBody] UpdateClassRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;

                var existingClass = await _classService.GetClassByIdAsync(id);
                if (existingClass == null)
                    return NotFound(new { message = "Không tìm thấy lớp" });

                // Kiểm tra quyền: Chỉ owner hoặc Admin mới được sửa
                if (userRole != UserRoles.Admin && existingClass.OwnerId != userId)
                {
                    return Forbid();
                }

                // Cập nhật các field
                if (!string.IsNullOrEmpty(request.Name))
                    existingClass.Name = request.Name;

                existingClass.Description = request.Description ?? existingClass.Description;
                existingClass.MaxStudents = request.MaxStudents ?? existingClass.MaxStudents;
                existingClass.AcademicYear = request.AcademicYear ?? existingClass.AcademicYear;
                existingClass.Grade = request.Grade ?? existingClass.Grade;

                if (request.IsActive.HasValue)
                    existingClass.IsActive = request.IsActive.Value;

                var updatedClass = await _classService.UpdateClassAsync(id, existingClass, userId);

                if (updatedClass == null)
                    return NotFound(new { message = "Không thể cập nhật lớp" });

                return Ok(new ClassResponse
                {
                    Id = updatedClass.Id ?? string.Empty,
                    Name = updatedClass.Name,
                    SchoolId = updatedClass.SchoolId,
                    OwnerId = updatedClass.OwnerId,
                    Description = updatedClass.Description,
                    MaxStudents = updatedClass.MaxStudents,
                    CurrentStudents = updatedClass.CurrentStudents,
                    AcademicYear = updatedClass.AcademicYear,
                    Grade = updatedClass.Grade,
                    StudentIds = updatedClass.StudentIds,
                    CreatedAt = updatedClass.CreatedAt,
                    IsActive = updatedClass.IsActive
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi cập nhật class");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }

        /// <summary>
        /// Xóa class
        /// Chỉ owner hoặc Admin mới được phép
        /// </summary>
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteClass(string id)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;

                var classEntity = await _classService.GetClassByIdAsync(id);
                if (classEntity == null)
                    return NotFound(new { message = "Không tìm thấy lớp" });

                // Kiểm tra quyền: Chỉ owner hoặc Admin mới được xóa
                if (userRole != UserRoles.Admin && classEntity.OwnerId != userId)
                {
                    return Forbid();
                }

                var result = await _classService.DeleteClassAsync(id);

                if (!result)
                    return BadRequest(new { message = "Không thể xóa lớp" });

                return Ok(new { message = "Đã xóa lớp thành công" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi xóa class");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }

        /// <summary>
        /// Thêm học sinh vào lớp
        /// </summary>
        [HttpPost("{classId}/students/{studentId}")]
        public async Task<IActionResult> AddStudentToClass(string classId, string studentId)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;

                var classEntity = await _classService.GetClassByIdAsync(classId);
                if (classEntity == null)
                    return NotFound(new { message = "Không tìm thấy lớp" });

                // Kiểm tra quyền
                if (userRole != UserRoles.Admin && classEntity.OwnerId != userId)
                {
                    return Forbid();
                }

                var result = await _classService.AddStudentToClassAsync(classId, studentId);

                if (!result)
                    return BadRequest(new { message = "Không thể thêm học sinh vào lớp" });

                return Ok(new { message = "Đã thêm học sinh vào lớp thành công" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi thêm học sinh vào lớp");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }

        /// <summary>
        /// Xóa học sinh khỏi lớp
        /// </summary>
        [HttpDelete("{classId}/students/{studentId}")]
        public async Task<IActionResult> RemoveStudentFromClass(string classId, string studentId)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;

                var classEntity = await _classService.GetClassByIdAsync(classId);
                if (classEntity == null)
                    return NotFound(new { message = "Không tìm thấy lớp" });

                // Kiểm tra quyền
                if (userRole != UserRoles.Admin && classEntity.OwnerId != userId)
                {
                    return Forbid();
                }

                var result = await _classService.RemoveStudentFromClassAsync(classId, studentId);

                if (!result)
                    return BadRequest(new { message = "Không thể xóa học sinh khỏi lớp" });

                return Ok(new { message = "Đã xóa học sinh khỏi lớp thành công" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi xóa học sinh khỏi lớp");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }
    }
}
