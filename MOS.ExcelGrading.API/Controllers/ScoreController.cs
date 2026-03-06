// MOS.ExcelGrading.API/Controllers/ScoreController.cs
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
    public class ScoreController : ControllerBase
    {
        private readonly IScoreService _scoreService;
        private readonly ILogger<ScoreController> _logger;

        public ScoreController(
            IScoreService scoreService,
            ILogger<ScoreController> logger)
        {
            _scoreService = scoreService;
            _logger = logger;
        }

        /// <summary>
        /// Lấy điểm theo bài tập
        /// </summary>
        [HttpGet("assignment/{assignmentId}")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> GetByAssignment(string assignmentId)
        {
            try
            {
                if (!HasPermission(Permissions.ViewGrades))
                    return Forbid();

                var scores = await _scoreService.GetScoresByAssignmentAsync(assignmentId);
                return Ok(scores);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting scores for assignment {AssignmentId}", assignmentId);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Lấy điểm theo học sinh
        /// </summary>
        [HttpGet("student/{studentId}")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> GetByStudent(string studentId)
        {
            try
            {
                if (!HasPermission(Permissions.ViewGrades))
                    return Forbid();

                var scores = await _scoreService.GetScoresByStudentAsync(studentId);
                return Ok(scores);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting scores for student {StudentId}", studentId);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Lấy báo cáo điểm của học sinh
        /// </summary>
        [HttpGet("student/{studentId}/class/{classId}/report")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> GetStudentReport(string studentId, string classId)
        {
            try
            {
                if (!HasPermission(Permissions.ViewGrades))
                    return Forbid();

                var report = await _scoreService.GetStudentScoreReportAsync(studentId, classId);
                return Ok(report);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting score report for student {StudentId}", studentId);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Chấm điểm cho 1 học sinh
        /// </summary>
        [HttpPost]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> CreateOrUpdate([FromBody] CreateScoreRequest request)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrEmpty(userId))
                    return Unauthorized();
                if (!HasPermission(Permissions.CreateGrades) && !HasPermission(Permissions.EditGrades))
                    return Forbid();

                var score = await _scoreService.CreateOrUpdateScoreAsync(request, userId);
                return Ok(score);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating/updating score");
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Chấm điểm hàng loạt
        /// </summary>
        [HttpPost("bulk")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> BulkCreateOrUpdate([FromBody] BulkScoreRequest request)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrEmpty(userId))
                    return Unauthorized();
                if (!HasPermission(Permissions.CreateGrades) && !HasPermission(Permissions.EditGrades))
                    return Forbid();

                var scores = await _scoreService.BulkCreateOrUpdateScoresAsync(request, userId);
                return Ok(new
                {
                    message = $"Chấm điểm thành công {scores.Count} học sinh",
                    scores = scores
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error bulk creating/updating scores");
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Xóa điểm
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
                if (!HasPermission(Permissions.DeleteGrades))
                    return Forbid();

                var result = await _scoreService.DeleteScoreAsync(id, userId);
                if (!result)
                    return NotFound(new { message = "Không tìm thấy điểm" });

                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting score {Id}", id);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        /// <summary>
        /// Lấy toàn bộ danh sách điểm của tất cả học sinh và tất cả bài tập trong một lớp
        /// </summary>
        [HttpGet("class/{classId}")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> GetByClass(string classId)
        {
            try
            {
                if (!HasPermission(Permissions.ViewGrades))
                    return Forbid();

                var scores = await _scoreService.GetScoresByClassAsync(classId);
                return Ok(scores);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting scores by class {ClassId}", classId);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        private bool HasPermission(string permission) =>
            User.Claims.Any(c => c.Type == "permission" && c.Value == permission);
    }
}

