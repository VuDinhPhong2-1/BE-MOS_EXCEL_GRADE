using System.Text;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/agent/exam-sessions")]
    [Authorize(Policy = "AgentApiAccess")]
    public class AgentExamSessionsController : ControllerBase
    {
        private static readonly Encoding Latin1 = Encoding.GetEncoding(1252);

        private readonly IExamSessionService _examSessionService;
        private readonly ILogger<AgentExamSessionsController> _logger;

        public AgentExamSessionsController(
            IExamSessionService examSessionService,
            ILogger<AgentExamSessionsController> logger)
        {
            _examSessionService = examSessionService;
            _logger = logger;
        }

        [HttpGet("{sessionId}/current-project-bootstrap")]
        public async Task<IActionResult> GetCurrentProjectBootstrap(string sessionId)
        {
            try
            {
                var bootstrap = await _examSessionService.GetCurrentProjectBootstrapAsync(sessionId);
                if (bootstrap == null)
                {
                    return NotFound(new { message = "Không tìm thấy project hiện tại" });
                }

                return Ok(bootstrap);
            }
            catch (ArgumentException ex)
            {
                var message = NormalizeMessage(ex.Message);
                _logger.LogWarning("Validation error getting current project bootstrap: {Message}", message);
                return BadRequest(new { message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting current project bootstrap");
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        [HttpPost("{sessionId}/projects/{projectCode}/score-upload")]
        public async Task<IActionResult> UploadScore(
            string sessionId,
            string projectCode,
            [FromBody] ScoreUploadRequest request)
        {
            try
            {
                await _examSessionService.UploadScoreAsync(sessionId, projectCode, request);
                return Ok(new { message = "Đã lưu kết quả chấm project" });
            }
            catch (ArgumentException ex)
            {
                var message = NormalizeMessage(ex.Message);
                _logger.LogWarning("Validation error uploading project score: {Message}", message);
                return BadRequest(new { message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error uploading project score");
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        [HttpPost("{sessionId}/advance")]
        public async Task<IActionResult> Advance(string sessionId)
        {
            try
            {
                var result = await _examSessionService.AdvanceAsync(sessionId);
                return Ok(result);
            }
            catch (ArgumentException ex)
            {
                var message = NormalizeMessage(ex.Message);
                _logger.LogWarning("Validation error advancing exam session: {Message}", message);
                return BadRequest(new { message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error advancing exam session");
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        [HttpPost("{sessionId}/restart-current-project")]
        public async Task<IActionResult> RestartCurrentProject(string sessionId)
        {
            try
            {
                var result = await _examSessionService.RestartCurrentProjectAsync(sessionId);
                return Ok(result);
            }
            catch (ArgumentException ex)
            {
                var message = NormalizeMessage(ex.Message);
                _logger.LogWarning("Validation error restarting current exam project: {Message}", message);
                return BadRequest(new { message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error restarting current exam project");
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        private static string NormalizeMessage(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return message;
            }

            if (!message.Contains('Ã') &&
                !message.Contains('Ä') &&
                !message.Contains('Æ') &&
                !message.Contains('á') &&
                !message.Contains('â'))
            {
                return message;
            }

            try
            {
                return Encoding.UTF8.GetString(Latin1.GetBytes(message));
            }
            catch
            {
                return message;
            }
        }
    }
}
