// MOS.ExcelGrading.API/Controllers/StudentController.cs
using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class StudentController : ControllerBase
    {
        private readonly IStudentService _studentService;
        private readonly ILogger<StudentController> _logger;

        public StudentController(
            IStudentService studentService,
            ILogger<StudentController> logger)
        {
            _studentService = studentService;
            _logger = logger;
        }

        // GET: api/student
        [HttpGet]
        [Authorize(Roles = $"{UserRoles.Admin},{UserRoles.Teacher}")]
        public async Task<IActionResult> GetAll()
        {
            try
            {
                var students = await _studentService.GetAllAsync();
                return Ok(students);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in GetAll");
                return StatusCode(500, new { message = "Internal server error" });
            }
        }

        // GET: api/student/{id}
        [HttpGet("{id}")]
        [Authorize(Roles = $"{UserRoles.Admin},{UserRoles.Teacher}")]
        public async Task<IActionResult> GetById(string id)
        {
            try
            {
                var student = await _studentService.GetByIdAsync(id);
                if (student == null)
                    return NotFound(new { message = "Student not found" });

                return Ok(student);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error in GetById: {id}");
                return StatusCode(500, new { message = "Internal server error" });
            }
        }

        // POST: api/student
        [HttpPost]
        [Authorize(Roles = $"{UserRoles.Admin},{UserRoles.Teacher}")]
        public async Task<IActionResult> Create([FromBody] CreateStudentRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "system";

                var student = await _studentService.CreateAsync(request, userId);

                return CreatedAtAction(
                    nameof(GetById),
                    new { id = student.Id },
                    student
                );
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in Create");
                return StatusCode(500, new { message = "Internal server error" });
            }
        }

        // PUT: api/student/{id}
        [HttpPut("{id}")]
        [Authorize(Roles = $"{UserRoles.Admin},{UserRoles.Teacher}")]
        public async Task<IActionResult> Update(string id, [FromBody] UpdateStudentRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "system";

                var student = await _studentService.UpdateAsync(id, request, userId);
                if (student == null)
                    return NotFound(new { message = "Student not found" });

                return Ok(student);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error in Update: {id}");
                return StatusCode(500, new { message = "Internal server error" });
            }
        }

        // DELETE: api/student/{id}
        [HttpDelete("{id}")]
        [Authorize(Roles = $"{UserRoles.Admin},{UserRoles.Teacher}")]
        public async Task<IActionResult> Delete(string id)
        {
            try
            {
                var result = await _studentService.DeleteAsync(id);
                if (!result)
                    return NotFound(new { message = "Student not found" });

                return Ok(new { message = "Student deleted successfully" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error in Delete: {id}");
                return StatusCode(500, new { message = "Internal server error" });
            }
        }
        // POST: api/student/import
        [HttpPost("import")]
        [Authorize(Roles = $"{UserRoles.Admin},{UserRoles.Teacher}")]
        public async Task<IActionResult> ImportFromExcel([FromForm] IFormFile file)
        {
            try
            {
                if (file == null || file.Length == 0)
                    return BadRequest(new { message = "File không hợp lệ" });

                // Kiểm tra extension
                var extension = Path.GetExtension(file.FileName).ToLower();
                if (extension != ".xlsx" && extension != ".xls")
                    return BadRequest(new { message = "Chỉ chấp nhận file Excel (.xlsx, .xls)" });

                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "system";

                // Optional: Lấy teacherId và teacherName từ form nếu có
                var teacherId = Request.Form["teacherId"].ToString();
                var teacherName = Request.Form["teacherName"].ToString();

                var result = await _studentService.ImportFromExcelAsync(
                    file,
                    userId,
                    string.IsNullOrEmpty(teacherId) ? null : teacherId,
                    string.IsNullOrEmpty(teacherName) ? null : teacherName
                );

                return Ok(new
                {
                    message = "Import hoàn tất",
                    data = result
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in ImportFromExcel");
                return StatusCode(500, new { message = "Internal server error" });
            }
        }

        // POST: api/student/bulk-import
        [HttpPost("bulk-import")]
        [Authorize(Roles = $"{UserRoles.Admin},{UserRoles.Teacher}")]
        public async Task<IActionResult> BulkImport([FromBody] BulkImportStudentRequest request)
        {
            try
            {
                if (request.Students == null || !request.Students.Any())
                    return BadRequest(new { message = "Danh sách học sinh trống" });

                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? "system";

                var result = await _studentService.BulkImportAsync(request, userId);

                return Ok(new
                {
                    message = $"Import hoàn tất: {result.SuccessCount}/{result.TotalCount} thành công",
                    data = result
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in BulkImport");
                return StatusCode(500, new { message = "Internal server error" });
            }
        }

        [HttpGet("class/{classId}")]
        [Authorize(Roles = $"{UserRoles.Admin},{UserRoles.Teacher}")]
        public async Task<IActionResult> GetByClassId(string classId)
        {
            try
            {
                var students = await _studentService.GetByClassIdAsync(classId);

                return Ok(new
                {
                    message = $"Tìm thấy {students.Count} học sinh",
                    data = students
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error in GetByClassId: {classId}");
                return StatusCode(500, new { message = "Internal server error" });
            }
        }


    }
}
