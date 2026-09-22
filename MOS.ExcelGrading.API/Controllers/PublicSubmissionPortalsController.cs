using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.Interfaces;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/public/portals")]
    public class PublicSubmissionPortalsController : ControllerBase
    {
        private readonly ISubmissionPortalService _service;
        private readonly ILogger<PublicSubmissionPortalsController> _logger;

        public PublicSubmissionPortalsController(ISubmissionPortalService service, ILogger<PublicSubmissionPortalsController> logger)
        {
            _service = service;
            _logger = logger;
        }

        [HttpGet("{token}")]
        public async Task<IActionResult> GetInfo(string token)
        {
            var info = await _service.GetPublicInfoAsync(token);
            return info == null ? NotFound(new { message = "Không tìm thấy link nộp bài" }) : Ok(info);
        }

        [HttpGet("{token}/classes/{classId}/students")]
        public async Task<IActionResult> GetStudents(string token, string classId)
        {
            try { return Ok(await _service.GetPublicStudentsAsync(token, classId)); }
            catch (InvalidOperationException ex) { return BadRequest(new { message = ex.Message }); }
        }

        [HttpPost("{token}/grade-and-submit")]
        [RequestSizeLimit(524288000)]
        public async Task<IActionResult> GradeAndSubmit(string token, [FromForm] string classId, [FromForm] string studentId, [FromForm] string assignmentId, [FromForm] IFormFile file)
        {
            try
            {
                var ip = HttpContext.Connection.RemoteIpAddress?.ToString();
                var userAgent = Request.Headers.UserAgent.ToString();
                return Ok(await _service.GradeAndSubmitAsync(token, classId, studentId, assignmentId, file, ip, userAgent));
            }
            catch (ArgumentException ex) { return BadRequest(new { message = ex.Message }); }
            catch (InvalidOperationException ex) { return BadRequest(new { message = ex.Message }); }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in public grade-and-submit");
                return StatusCode(500, new { message = "Lỗi máy chủ nội bộ" });
            }
        }

        [HttpGet("{token}/leaderboard")]
        public async Task<IActionResult> GetLeaderboard(string token, [FromQuery] string? classId = null, [FromQuery] string? assignmentId = null)
        {
            try { return Ok(await _service.GetLeaderboardAsync(token, classId, assignmentId)); }
            catch (InvalidOperationException ex) { return BadRequest(new { message = ex.Message }); }
        }
    }
}