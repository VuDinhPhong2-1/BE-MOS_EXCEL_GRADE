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
        private readonly IUserService _userService;
        private readonly ILogger<ClassController> _logger;

        public ClassController(
            IClassService classService,
            ISchoolService schoolService,
            IUserService userService,
            ILogger<ClassController> logger)
        {
            _classService = classService;
            _schoolService = schoolService;
            _userService = userService;
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

                var response = classes.Select(ToClassResponse).ToList();

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
        /// Teacher/Admin: Được xem tất cả class trong trường
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
                    return BadRequest(new { message = "Mã trường không hợp lệ" });
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

                var response = classes.Select(ToClassResponse).ToList();

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
                var classEntity = await _classService.GetClassByIdAsync(id);

                if (classEntity == null)
                    return NotFound(new { message = "Không tìm thấy lớp" });

                var response = ToClassResponse(classEntity);

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

                // Kiểm tra tên lớp đã tồn tại trong school hiện tại chưa
                if (await _classService.ClassExistsAsync(request.SchoolId, request.Name))
                    return BadRequest(new { message = "Tên lớp đã tồn tại trong trường này" });

                var classEntity = new Class
                {
                    Name = request.Name,
                    SchoolId = request.SchoolId,
                    Description = request.Description,
                    MaxStudents = request.MaxStudents,
                    AcademicYear = request.AcademicYear,
                    Grade = request.Grade,
                    AttendanceSpreadsheetId = NormalizeOptional(request.AttendanceSpreadsheetId),
                    AttendanceWorksheetName = NormalizeOptional(request.AttendanceWorksheetName),
                };

                var createdClass = await _classService.CreateClassAsync(classEntity, userId);

                return CreatedAtAction(
                    nameof(GetClassById),
                    new { id = createdClass.Id },
                    ToClassResponse(createdClass)
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
                if (!CanManageClass(existingClass, userId, userRole))
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
                if (request.AttendanceSpreadsheetId != null)
                    existingClass.AttendanceSpreadsheetId = NormalizeOptional(request.AttendanceSpreadsheetId);
                if (request.AttendanceWorksheetName != null)
                    existingClass.AttendanceWorksheetName = NormalizeOptional(request.AttendanceWorksheetName);

                if (request.IsActive.HasValue)
                    existingClass.IsActive = request.IsActive.Value;

                var updatedClass = await _classService.UpdateClassAsync(id, existingClass, userId);

                if (updatedClass == null)
                    return NotFound(new { message = "Không thể cập nhật lớp" });

                return Ok(ToClassResponse(updatedClass));
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
                if (!CanManageClass(classEntity, userId, userRole))
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
                if (!CanManageClass(classEntity, userId, userRole))
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
                if (!CanManageClass(classEntity, userId, userRole))
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

        [HttpPost("{id}/handover")]
        public async Task<IActionResult> GrantClassManagement(string id, [FromBody] ClassHandoverRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
                var teacherId = request.TeacherId?.Trim() ?? string.Empty;

                if (string.IsNullOrWhiteSpace(teacherId))
                    return BadRequest(new { message = "TeacherId là bắt buộc" });

                var classEntity = await _classService.GetClassByIdAsync(id);
                if (classEntity == null)
                    return NotFound(new { message = "Không tìm thấy lớp" });

                if (userRole != UserRoles.Admin && classEntity.OwnerId != userId)
                    return Forbid();

                var teacher = await _userService.GetUserByIdAsync(teacherId);
                if (teacher == null || teacher.Role != UserRoles.Teacher || !teacher.IsActive)
                    return BadRequest(new { message = "Giáo viên không hợp lệ hoặc đã bị vô hiệu hóa" });

                if (teacherId == classEntity.OwnerId)
                    return BadRequest(new { message = "Giáo viên chính đã có toàn quyền lớp này" });

                classEntity.ManagerTeacherIds ??= new List<string>();

                if (!classEntity.ManagerTeacherIds.Contains(teacherId))
                    classEntity.ManagerTeacherIds.Add(teacherId);

                var updatedClass = await _classService.UpdateClassAsync(id, classEntity, userId);
                if (updatedClass == null)
                    return NotFound(new { message = "Không thể cập nhật bàn giao lớp" });

                return Ok(ToClassResponse(updatedClass));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi bàn giao quyền quản lý lớp");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }

        [HttpDelete("{id}/handover/{teacherId}")]
        public async Task<IActionResult> RevokeClassManagement(string id, string teacherId)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var userRole = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
                var targetTeacherId = teacherId?.Trim() ?? string.Empty;

                if (string.IsNullOrWhiteSpace(targetTeacherId))
                    return BadRequest(new { message = "TeacherId là bắt buộc" });

                var classEntity = await _classService.GetClassByIdAsync(id);
                if (classEntity == null)
                    return NotFound(new { message = "Không tìm thấy lớp" });

                if (userRole != UserRoles.Admin && classEntity.OwnerId != userId)
                    return Forbid();

                if (targetTeacherId == classEntity.OwnerId)
                    return BadRequest(new { message = "Không thể thu hồi quyền của giáo viên chính" });

                classEntity.ManagerTeacherIds ??= new List<string>();
                classEntity.ManagerTeacherIds.RemoveAll(x => x == targetTeacherId);

                var updatedClass = await _classService.UpdateClassAsync(id, classEntity, userId);
                if (updatedClass == null)
                    return NotFound(new { message = "Không thể cập nhật bàn giao lớp" });

                return Ok(ToClassResponse(updatedClass));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi thu hồi quyền quản lý lớp");
                return StatusCode(500, new { message = "Đã xảy ra lỗi" });
            }
        }

        private static bool CanManageClass(Class classEntity, string userId, string userRole)
        {
            if (userRole == UserRoles.Admin)
                return true;

            if (classEntity.OwnerId == userId)
                return true;

            return classEntity.ManagerTeacherIds?.Contains(userId) == true;
        }

        private static ClassResponse ToClassResponse(Class classEntity)
        {
            return new ClassResponse
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
                AttendanceSpreadsheetId = classEntity.AttendanceSpreadsheetId,
                AttendanceWorksheetName = classEntity.AttendanceWorksheetName,
                StudentIds = classEntity.StudentIds,
                ManagerTeacherIds = classEntity.ManagerTeacherIds ?? new List<string>(),
                CreatedAt = classEntity.CreatedAt,
                IsActive = classEntity.IsActive
            };
        }

        private static string? NormalizeOptional(string? value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? null
                : value.Trim();
        }
    }
}


