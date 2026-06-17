using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/exam-publications")]
    [Authorize]
    public class ExamPublicationsController : ControllerBase
    {
        private readonly IExamPublicationService _examPublicationService;
        private readonly ILogger<ExamPublicationsController> _logger;

        public ExamPublicationsController(
            IExamPublicationService examPublicationService,
            ILogger<ExamPublicationsController> logger)
        {
            _examPublicationService = examPublicationService;
            _logger = logger;
        }

        [HttpGet("{id}")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> GetById(string id)
        {
            try
            {
                if (!HasPermission(Permissions.ViewProjects))
                {
                    return Forbid();
                }

                var publication = await _examPublicationService.GetExamPublicationByIdAsync(id);
                if (publication == null)
                {
                    return NotFound(new { message = "Không tìm thấy ca thi" });
                }

                return Ok(publication);
            }
            catch (ArgumentException ex)
            {
                _logger.LogWarning(ex, "Validation error getting exam publication {PublicationId}", id);
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting exam publication {PublicationId}", id);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        [HttpPost]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> Create([FromBody] CreateExamPublicationRequest request)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrWhiteSpace(userId))
                {
                    return Unauthorized();
                }

                if (!HasPermission(Permissions.CreateProjects))
                {
                    return Forbid();
                }

                var publication = await _examPublicationService.CreateExamPublicationAsync(request, userId);
                return CreatedAtAction(nameof(GetById), new { id = publication.Id }, publication);
            }
            catch (ArgumentException ex)
            {
                _logger.LogWarning(ex, "Validation error creating exam publication");
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating exam publication");
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        private bool HasPermission(string permission) =>
            User.Claims.Any(c => c.Type == "permission" && c.Value == permission);
    }
}
