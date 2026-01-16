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
        private readonly ILogger<AssignmentController> _logger;

        public AssignmentController(
            IAssignmentService assignmentService,
            ILogger<AssignmentController> logger)
        {
            _assignmentService = assignmentService;
            _logger = logger;
        }
        /// <summary>
        /// Lấy danh sách các Grading API endpoints có sẵn
        /// </summary>
        [HttpGet("grading-endpoints")]
        public IActionResult GetGradingEndpoints()
        {
            var endpoints = new List<GradingEndpointInfo>
    {
        new() {
            Endpoint = GradingApiEndpoints.Project09,
            DisplayName = "Project 09",
            Description = "Chấm điểm Project 09",
            MaxScore = 32
        },
        new() {
            Endpoint = GradingApiEndpoints.Project10,
            DisplayName = "Project 10",
            Description = "Chấm điểm Project 10",
            MaxScore = 25 // <-- Gán giá trị tạm cho project này
        },
        new() {
            Endpoint = GradingApiEndpoints.Project11,
            DisplayName = "Project 11",
            Description = "Chấm điểm Project 11",
            MaxScore = 30 // <-- Gán giá trị tạm
        },
        // ... các project khác
        new() {
            Endpoint = GradingApiEndpoints.Project16,
            DisplayName = "Project 16",
            Description = "Chấm điểm Project 16",
            MaxScore = 28 // <-- Gán giá trị tạm
        }
    };

            return Ok(endpoints);
        }
        
        /// <summary>
        /// Lấy danh sách bài tập theo lớp
        /// </summary>
        [HttpGet("class/{classId}")]
        public async Task<IActionResult> GetByClass(string classId)
        {
            try
            {
                var assignments = await _assignmentService.GetAssignmentsByClassIdAsync(classId);
                return Ok(assignments);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting assignments for class {ClassId}", classId);
                return StatusCode(500, "Internal server error");
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
                return StatusCode(500, "Internal server error");
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
                    return NotFound(new { message = "Assignment not found" });

                return Ok(assignment);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting assignment {Id}", id);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Tạo bài tập mới
        /// </summary>
        [HttpPost]
        public async Task<IActionResult> Create([FromBody] CreateAssignmentRequest request)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrEmpty(userId))
                    return Unauthorized();

                var assignment = await _assignmentService.CreateAssignmentAsync(request, userId);
                return CreatedAtAction(nameof(GetById), new { id = assignment.Id }, assignment);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating assignment");
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Cập nhật bài tập
        /// </summary>
        [HttpPut("{id}")]
        public async Task<IActionResult> Update(string id, [FromBody] UpdateAssignmentRequest request)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrEmpty(userId))
                    return Unauthorized();

                var assignment = await _assignmentService.UpdateAssignmentAsync(id, request, userId);
                if (assignment == null)
                    return NotFound(new { message = "Assignment not found" });

                return Ok(assignment);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating assignment {Id}", id);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Xóa bài tập (soft delete)
        /// </summary>
        [HttpDelete("{id}")]
        public async Task<IActionResult> Delete(string id)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrEmpty(userId))
                    return Unauthorized();

                var result = await _assignmentService.DeleteAssignmentAsync(id, userId);
                if (!result)
                    return NotFound(new { message = "Assignment not found" });

                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting assignment {Id}", id);
                return StatusCode(500, "Internal server error");
            }
        }
    }
}
