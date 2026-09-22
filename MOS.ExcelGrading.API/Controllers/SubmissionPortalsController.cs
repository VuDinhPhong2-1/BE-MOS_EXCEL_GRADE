using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/submission-portals")]
    [Authorize]
    public class SubmissionPortalsController : ControllerBase
    {
        private readonly ISubmissionPortalService _service;
        private readonly ILogger<SubmissionPortalsController> _logger;

        public SubmissionPortalsController(ISubmissionPortalService service, ILogger<SubmissionPortalsController> logger)
        {
            _service = service;
            _logger = logger;
        }

        [HttpGet]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> GetAll()
        {
            if (!HasPermission(Permissions.ViewProjects)) return Forbid();
            var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
            var result = await _service.GetAllAsync(userId, User.IsInRole(UserRoles.Admin));
            return Ok(result);
        }

        [HttpGet("{id}")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> GetById(string id)
        {
            if (!HasPermission(Permissions.ViewProjects)) return Forbid();
            var portal = await _service.GetByIdAsync(id);
            return portal == null ? NotFound(new { message = "Không tìm thấy link nộp bài" }) : Ok(portal);
        }

        [HttpPost]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> Create([FromBody] CreateSubmissionPortalRequest request)
        {
            try
            {
                if (!HasPermission(Permissions.CreateProjects)) return Forbid();
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrWhiteSpace(userId)) return Unauthorized();
                var portal = await _service.CreateAsync(request, userId);
                return CreatedAtAction(nameof(GetById), new { id = portal.Id }, portal);
            }
            catch (ArgumentException ex) { return BadRequest(new { message = ex.Message }); }
            catch (InvalidOperationException ex) { return BadRequest(new { message = ex.Message }); }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating submission portal");
                return StatusCode(500, new { message = "Lỗi máy chủ nội bộ" });
            }
        }

        [HttpPut("{id}")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> Update(string id, [FromBody] UpdateSubmissionPortalRequest request)
        {
            try
            {
                if (!HasPermission(Permissions.EditProjects)) return Forbid();
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var portal = await _service.UpdateAsync(id, request, userId);
                return portal == null ? NotFound(new { message = "Không tìm thấy link nộp bài" }) : Ok(portal);
            }
            catch (ArgumentException ex) { return BadRequest(new { message = ex.Message }); }
            catch (InvalidOperationException ex) { return BadRequest(new { message = ex.Message }); }
        }

        [HttpDelete("{id}")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> Delete(string id)
        {
            if (!HasPermission(Permissions.DeleteProjects)) return Forbid();
            var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
            var ok = await _service.DeleteAsync(id, userId);
            return ok ? NoContent() : NotFound(new { message = "Không tìm thấy link nộp bài" });
        }

        [HttpGet("{id}/alerts")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> GetAlerts(string id, [FromQuery] bool includeDismissed = false)
        {
            if (!HasPermission(Permissions.ViewProjects)) return Forbid();
            return Ok(await _service.GetAlertsAsync(id, includeDismissed));
        }

        [HttpGet("{id}/alerts/summary")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> GetAlertSummary(string id)
        {
            if (!HasPermission(Permissions.ViewProjects)) return Forbid();
            return Ok(new { unreadAlertCount = await _service.GetUnreadAlertCountAsync(id) });
        }

        [HttpPut("{id}/alerts/{alertId}")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> UpdateAlert(string id, string alertId, [FromBody] UpdateSubmissionAlertRequest request)
        {
            if (!HasPermission(Permissions.EditProjects)) return Forbid();
            var ok = await _service.UpdateAlertAsync(id, alertId, request.IsRead, request.IsDismissed);
            return ok ? Ok(new { message = "Đã cập nhật cảnh báo" }) : NotFound(new { message = "Không tìm thấy cảnh báo" });
        }

        [HttpGet("{id}/submission-logs")]
        [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
        public async Task<IActionResult> GetLogs(string id)
        {
            if (!HasPermission(Permissions.ViewGrades)) return Forbid();
            return Ok(await _service.GetSubmissionLogsAsync(id));
        }

        private bool HasPermission(string permission) => User.Claims.Any(c => c.Type == "permission" && c.Value == permission);
    }

    public class UpdateSubmissionAlertRequest
    {
        public bool? IsRead { get; set; }
        public bool? IsDismissed { get; set; }
    }
}