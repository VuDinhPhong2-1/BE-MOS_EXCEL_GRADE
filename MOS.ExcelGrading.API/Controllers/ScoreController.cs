// MOS.ExcelGrading.API/Controllers/ScoreController.cs
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
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
        public async Task<IActionResult> GetByAssignment(string assignmentId)
        {
            try
            {
                var scores = await _scoreService.GetScoresByAssignmentAsync(assignmentId);
                return Ok(scores);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting scores for assignment {AssignmentId}", assignmentId);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Lấy điểm theo học sinh
        /// </summary>
        [HttpGet("student/{studentId}")]
        public async Task<IActionResult> GetByStudent(string studentId)
        {
            try
            {
                var scores = await _scoreService.GetScoresByStudentAsync(studentId);
                return Ok(scores);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting scores for student {StudentId}", studentId);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Lấy báo cáo điểm của học sinh
        /// </summary>
        [HttpGet("student/{studentId}/class/{classId}/report")]
        public async Task<IActionResult> GetStudentReport(string studentId, string classId)
        {
            try
            {
                var report = await _scoreService.GetStudentScoreReportAsync(studentId, classId);
                return Ok(report);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting score report for student {StudentId}", studentId);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Chấm điểm cho 1 học sinh
        /// </summary>
        [HttpPost]
        public async Task<IActionResult> CreateOrUpdate([FromBody] CreateScoreRequest request)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrEmpty(userId))
                    return Unauthorized();

                var score = await _scoreService.CreateOrUpdateScoreAsync(request, userId);
                return Ok(score);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating/updating score");
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Chấm điểm hàng loạt
        /// </summary>
        [HttpPost("bulk")]
        public async Task<IActionResult> BulkCreateOrUpdate([FromBody] BulkScoreRequest request)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrEmpty(userId))
                    return Unauthorized();

                var scores = await _scoreService.BulkCreateOrUpdateScoresAsync(request, userId);
                return Ok(new
                {
                    message = $"Successfully graded {scores.Count} students",
                    scores = scores
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error bulk creating/updating scores");
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Xóa điểm
        /// </summary>
        [HttpDelete("{id}")]
        public async Task<IActionResult> Delete(string id)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrEmpty(userId))
                    return Unauthorized();

                var result = await _scoreService.DeleteScoreAsync(id, userId);
                if (!result)
                    return NotFound(new { message = "Score not found" });

                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting score {Id}", id);
                return StatusCode(500, "Internal server error");
            }
        }
    }
}
