// MOS.ExcelGrading.API/Controllers/AssignmentController.cs
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
    public class AssignmentController : ControllerBase
    {
        private readonly IAssignmentService _assignmentService;
        private readonly IAssignmentFileService _assignmentFileService;
        private readonly ILogger<AssignmentController> _logger;

        public AssignmentController(
            IAssignmentService assignmentService,
            IAssignmentFileService assignmentFileService,
            ILogger<AssignmentController> logger)
        {
            _assignmentService = assignmentService;
            _assignmentFileService = assignmentFileService;
            _logger = logger;
        }
        /// <summary>
        /// Lấy danh sách các Grading API endpoints có sẵn
        /// </summary>
        [HttpGet("grading-endpoints")]
        public IActionResult GetGradingEndpoints()
        {
            var implementedExcelProjects = new List<(int Number, string Description, double RawMaxScore)>
            {
                (1, "Chấm điểm Dự án 01", 125),
                (2, "Chấm điểm Dự án 02", 125),
                (3, "Chấm điểm Dự án 03 (Task 1-5 tự động, Task 6 thủ công)", 125),
                (4, "Chấm điểm Dự án 04", 125),
                (5, "Chấm điểm Dự án 05", 125),
                (6, "Chấm điểm Dự án 06", 125),
                (7, "Chấm điểm Dự án 07", 125),
                (8, "Chấm điểm Dự án 08", 125),
                (9, "Chấm điểm Dự án 09", 125),
                (10, "Chấm điểm Dự án 10", 125),
                (11, "Chấm điểm Dự án 11", 125),
                (12, "Chấm điểm Dự án 12", 125),
                (13, "Chấm điểm Dự án 13", 125),
                (14, "Chấm điểm Dự án 14", 125),
                (15, "Chấm điểm Dự án 15", 125),
                (16, "Chấm điểm Dự án 16", 125),
                (18, "Chấm điểm Dự án 18", 125),
                (20, "Chấm điểm Dự án 20", 125),
                (22, "Chấm điểm Dự án 22", 125)
            };

            var endpoints = implementedExcelProjects
                .Select(item => BuildExcelEndpointInfo(item.Number, item.Description, item.RawMaxScore))
                .OrderBy(item => item.DisplayName, StringComparer.Create(System.Globalization.CultureInfo.GetCultureInfo("vi-VN"), ignoreCase: true))
                .ToList();

            return Ok(endpoints);
        }

        private static GradingEndpointInfo BuildExcelEndpointInfo(int projectNumber, string baseDescription, double rawMaxScore)
        {
            var endpoint = GradingApiEndpoints.ToExcelProjectEndpoint(projectNumber);
            var practice = PracticeScoring.ResolveByProjectNumber(projectNumber);
            var practiceProjectScore = (double)PracticeScoring.CalculateProjectMaxScore(projectNumber);

            return new GradingEndpointInfo
            {
                Endpoint = endpoint,
                DisplayName = $"Project {projectNumber:00} - Excel",
                Description =
                    $"{baseDescription}. Quy đổi theo {practice.Name}: {practiceProjectScore:0.##}/{practice.TotalScore} điểm.",
                MaxScore = practiceProjectScore,
                RawMaxScore = rawMaxScore,
                Subject = GradingApiSubjects.Excel,
                PracticeCode = practice.Code,
                PracticeName = practice.Name,
                PracticeTotalScore = practice.TotalScore,
                PracticeProjectCount = practice.ProjectCount,
                ApiPath = $"/api/grading/{endpoint}"
            };
        }
        
        /// <summary>
        /// Lấy danh sách bài tập theo lớp
        /// </summary>
        [HttpGet("class/{classId}")]
        public async Task<IActionResult> GetByClass(string classId, [FromQuery] bool includeInactive = false)
        {
            try
            {
                var assignments = await _assignmentService.GetAssignmentsByClassIdAsync(classId, includeInactive);
                return Ok(assignments);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting assignments for class {ClassId}", classId);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Lấy danh sách bài tập kèm thống kê
        /// </summary>
        [HttpGet("class/{classId}/stats")]
        public async Task<IActionResult> GetByClassWithStats(string classId)
        {
            try
            {
                var assignments = await _assignmentService.GetAssignmentsWithStatsByClassIdAsync(classId);
                return Ok(assignments);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting assignments with stats for class {ClassId}", classId);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Lấy chi tiết bài tập
        /// </summary>
        [HttpGet("{id}")]
        public async Task<IActionResult> GetById(string id)
        {
            try
            {
                var assignment = await _assignmentService.GetAssignmentByIdAsync(id);
                if (assignment == null)
                    return NotFound(new { message = "Không tìm thấy bài tập" });

                return Ok(assignment);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting assignment {Id}", id);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Upload file bài tập (template/answer) cho assignment
        /// </summary>
        [HttpPost("{assignmentId}/files")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> UploadAssignmentFile(
            string assignmentId,
            [FromForm] IFormFile file,
            [FromForm] string subject,
            [FromForm] string kind)
        {
            try
            {
                var userId = GetCurrentUserId();
                if (userId == null)
                {
                    return Unauthorized();
                }

                if (!HasPermission(Permissions.EditProjects) && !HasPermission(Permissions.CreateProjects))
                {
                    return Forbid();
                }

                var uploaded = await _assignmentFileService.UploadAssignmentFileAsync(
                    assignmentId,
                    file,
                    subject,
                    kind,
                    userId);

                return Ok(uploaded);
            }
            catch (ArgumentException ex)
            {
                _logger.LogWarning(ex, "Validation error while uploading assignment file");
                return BadRequest(new { message = ex.Message });
            }
            catch (InvalidOperationException ex)
            {
                _logger.LogWarning(ex, "Not found while uploading assignment file for assignment {AssignmentId}", assignmentId);
                return NotFound(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error uploading assignment file for assignment {AssignmentId}", assignmentId);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Lấy danh sách file đã upload của assignment
        /// </summary>
        [HttpGet("{assignmentId}/files")]
        public async Task<IActionResult> GetAssignmentFiles(
            string assignmentId,
            [FromQuery] bool includeInactive = false)
        {
            try
            {
                if (!HasPermission(Permissions.ViewProjects))
                {
                    return Forbid();
                }

                var files = await _assignmentFileService.GetAssignmentFilesAsync(assignmentId, includeInactive);
                return Ok(files);
            }
            catch (ArgumentException ex)
            {
                _logger.LogWarning(ex, "Validation error while getting assignment files");
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting assignment files for assignment {AssignmentId}", assignmentId);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Lấy metadata chi tiết của file bài tập
        /// </summary>
        [HttpGet("files/{fileId}")]
        public async Task<IActionResult> GetAssignmentFileById(string fileId)
        {
            try
            {
                if (!HasPermission(Permissions.ViewProjects))
                {
                    return Forbid();
                }

                var file = await _assignmentFileService.GetAssignmentFileByIdAsync(fileId);
                if (file == null)
                {
                    return NotFound(new { message = "Không tìm thấy file bài tập" });
                }

                return Ok(file);
            }
            catch (ArgumentException ex)
            {
                _logger.LogWarning(ex, "Validation error while getting assignment file {FileId}", fileId);
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting assignment file {FileId}", fileId);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Download file bài tập theo fileId
        /// </summary>
        [HttpGet("files/{fileId}/download")]
        public async Task<IActionResult> DownloadAssignmentFile(string fileId)
        {
            try
            {
                if (!HasPermission(Permissions.ViewProjects))
                {
                    return Forbid();
                }

                var fileResult = await _assignmentFileService.OpenAssignmentFileDownloadAsync(fileId);
                return File(fileResult.Stream, fileResult.ContentType, fileResult.FileName);
            }
            catch (ArgumentException ex)
            {
                _logger.LogWarning(ex, "Validation error while downloading assignment file {FileId}", fileId);
                return BadRequest(new { message = ex.Message });
            }
            catch (InvalidOperationException ex)
            {
                _logger.LogWarning(ex, "Not found while downloading assignment file {FileId}", fileId);
                return NotFound(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error downloading assignment file {FileId}", fileId);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Xóa mềm file bài tập
        /// </summary>
        [HttpDelete("files/{fileId}")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> DeleteAssignmentFile(string fileId)
        {
            try
            {
                var userId = GetCurrentUserId();
                if (userId == null)
                {
                    return Unauthorized();
                }

                if (!HasPermission(Permissions.DeleteProjects) && !HasPermission(Permissions.EditProjects))
                {
                    return Forbid();
                }

                var deleted = await _assignmentFileService.SoftDeleteAssignmentFileAsync(fileId, userId);
                if (!deleted)
                {
                    return NotFound(new { message = "Không tìm thấy file bài tập" });
                }

                return NoContent();
            }
            catch (ArgumentException ex)
            {
                _logger.LogWarning(ex, "Validation error while deleting assignment file {FileId}", fileId);
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting assignment file {FileId}", fileId);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Tạo bài tập mới
        /// </summary>
        [HttpPost]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> Create([FromBody] CreateAssignmentRequest request)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrEmpty(userId))
                    return Unauthorized();
                if (!HasPermission(Permissions.CreateProjects))
                    return Forbid();

                var assignment = await _assignmentService.CreateAssignmentAsync(request, userId);
                return CreatedAtAction(nameof(GetById), new { id = assignment.Id }, assignment);
            }
            catch (ArgumentException ex)
            {
                _logger.LogWarning(ex, "Validation error creating assignment");
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating assignment");
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Cập nhật bài tập
        /// </summary>
        [HttpPut("{id}")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> Update(string id, [FromBody] UpdateAssignmentRequest request)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrEmpty(userId))
                    return Unauthorized();
                if (!HasPermission(Permissions.EditProjects))
                    return Forbid();

                var assignment = await _assignmentService.UpdateAssignmentAsync(id, request, userId);
                if (assignment == null)
                    return NotFound(new { message = "Không tìm thấy bài tập" });

                return Ok(assignment);
            }
            catch (ArgumentException ex)
            {
                _logger.LogWarning(ex, "Validation error updating assignment {Id}", id);
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating assignment {Id}", id);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Xóa bài tập (soft delete)
        /// </summary>
        [HttpDelete("{id}")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> Delete(string id)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrEmpty(userId))
                    return Unauthorized();
                if (!HasPermission(Permissions.DeleteProjects))
                    return Forbid();

                var result = await _assignmentService.DeleteAssignmentAsync(id, userId);
                if (!result)
                    return NotFound(new { message = "Không tìm thấy bài tập" });

                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting assignment {Id}", id);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        private bool HasPermission(string permission) =>
            User.Claims.Any(c => c.Type == "permission" && c.Value == permission);

        private string? GetCurrentUserId() =>
            User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
    }
}

