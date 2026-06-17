using System.Text;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/public/exams/{token}/sessions")]
    public class PublicExamSessionsController : ControllerBase
    {
        private static readonly Encoding Latin1 = Encoding.GetEncoding(1252);

        private readonly IExamSessionService _examSessionService;
        private readonly ILogger<PublicExamSessionsController> _logger;

        public PublicExamSessionsController(
            IExamSessionService examSessionService,
            ILogger<PublicExamSessionsController> logger)
        {
            _examSessionService = examSessionService;
            _logger = logger;
        }

        [HttpPost]
        public async Task<IActionResult> StartSession(
            string token,
            [FromBody] StartExamSessionRequest request)
        {
            try
            {
                var session = await _examSessionService.StartSessionAsync(token, request);
                var state = await _examSessionService.GetStateAsync(token, session.Id);

                return Ok(state);
            }
            catch (ArgumentException ex)
            {
                var message = NormalizeMessage(ex.Message);
                _logger.LogWarning("Validation error starting exam session: {Message}", message);
                return BadRequest(new { message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error starting exam session");
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        [HttpGet("{sessionId}")]
        public async Task<IActionResult> GetState(string token, string sessionId)
        {
            try
            {
                var state = await _examSessionService.GetStateAsync(token, sessionId);
                if (state == null)
                {
                    return NotFound(new { message = "Không tìm thấy phiên thi" });
                }

                return Ok(state);
            }
            catch (ArgumentException ex)
            {
                var message = NormalizeMessage(ex.Message);
                _logger.LogWarning("Validation error getting exam session state: {Message}", message);
                return BadRequest(new { message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting exam session state");
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
