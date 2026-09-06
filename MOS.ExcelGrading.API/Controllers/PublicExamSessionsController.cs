using System.Text;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/public/exams/{token}/sessions")]
    public class PublicExamSessionsController : ControllerBase
    {
        private static readonly Encoding Latin1 = Encoding.GetEncoding(1252);

        private readonly IExamSessionService _examSessionService;
        private readonly IXmlGradingRuleService _xmlGradingRuleService;
        private readonly ILogger<PublicExamSessionsController> _logger;

        public PublicExamSessionsController(
            IExamSessionService examSessionService,
            IXmlGradingRuleService xmlGradingRuleService,
            ILogger<PublicExamSessionsController> logger)
        {
            _examSessionService = examSessionService;
            _xmlGradingRuleService = xmlGradingRuleService;
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
                return StatusCode(500, "Loi may chu noi bo");
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
                    return NotFound(new { message = "Khong tim thay phien thi" });
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
                return StatusCode(500, "Loi may chu noi bo");
            }
        }

        [HttpGet("{sessionId}/current-project-bootstrap")]
        public async Task<IActionResult> GetCurrentProjectBootstrap(string token, string sessionId)
        {
            try
            {
                var bootstrap = await _examSessionService.GetCurrentProjectBootstrapAsync(token, sessionId);
                if (bootstrap == null)
                {
                    return NotFound(new { message = "Khong tim thay project hien tai" });
                }

                return Ok(bootstrap);
            }
            catch (ArgumentException ex)
            {
                var message = NormalizeMessage(ex.Message);
                _logger.LogWarning("Validation error getting public project bootstrap: {Message}", message);
                return BadRequest(new { message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting public project bootstrap");
                return StatusCode(500, "Loi may chu noi bo");
            }
        }

        [HttpPost("{sessionId}/restart-current-project")]
        public async Task<IActionResult> RestartCurrentProject(string token, string sessionId)
        {
            try
            {
                var result = await _examSessionService.RestartCurrentProjectAsync(token, sessionId);
                return Ok(result);
            }
            catch (ArgumentException ex)
            {
                var message = NormalizeMessage(ex.Message);
                _logger.LogWarning("Validation error restarting public exam project: {Message}", message);
                return BadRequest(new { message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error restarting public exam project");
                return StatusCode(500, "Loi may chu noi bo");
            }
        }

        [HttpPost("{sessionId}/advance")]
        public async Task<IActionResult> Advance(string token, string sessionId)
        {
            try
            {
                var result = await _examSessionService.AdvanceAsync(token, sessionId);
                return Ok(result);
            }
            catch (ArgumentException ex)
            {
                var message = NormalizeMessage(ex.Message);
                _logger.LogWarning("Validation error advancing public exam session: {Message}", message);
                return BadRequest(new { message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error advancing public exam session");
                return StatusCode(500, "Loi may chu noi bo");
            }
        }

        [HttpPost("{sessionId}/projects/{projectCode}/grade")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeCurrentProject(
            string token,
            string sessionId,
            string projectCode,
            [FromForm] IFormFile studentFile)
        {
            try
            {
                if (studentFile == null)
                {
                    return BadRequest(new { message = "Can cung cap file: studentFile" });
                }

                var bootstrap = await _examSessionService.GetCurrentProjectBootstrapAsync(token, sessionId);
                if (bootstrap == null)
                {
                    return NotFound(new { message = "Khong tim thay project hien tai" });
                }

                if (!string.Equals(bootstrap.ProjectCode, projectCode, StringComparison.OrdinalIgnoreCase))
                {
                    return BadRequest(new { message = "ProjectCode khong khop voi project hien tai." });
                }

                var result = await GradeByBootstrapAsync(bootstrap, studentFile);

                await _examSessionService.UploadScoreAsync(
                    token,
                    sessionId,
                    projectCode,
                    new ScoreUploadRequest
                    {
                        WorkingFilePath = Path.GetFileName(studentFile.FileName),
                        Score = (double)result.TotalScore,
                        MaxScore = (double)result.MaxScore,
                        Feedback = result.Status,
                        IsSuccess = true
                    });

                return Ok(result);
            }
            catch (ArgumentException ex)
            {
                var message = NormalizeMessage(ex.Message);
                _logger.LogWarning("Validation error grading public exam project: {Message}", message);
                return BadRequest(new { message });
            }
            catch (InvalidOperationException ex)
            {
                var message = NormalizeMessage(ex.Message);
                _logger.LogWarning("Business error grading public exam project: {Message}", message);
                return BadRequest(new { message });
            }
            catch (InvalidDataException ex)
            {
                var message = NormalizeMessage(ex.Message);
                _logger.LogWarning("Invalid Office package uploaded for public exam grading: {Message}", message);
                return BadRequest(new { message = "File upload không phải Office OpenXML hợp lệ." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading public exam project");
                return StatusCode(500, "Loi may chu noi bo");
            }
        }

        private async Task<GradingResult> GradeByBootstrapAsync(
            ExamSessionProjectBootstrapDto bootstrap,
            IFormFile studentFile)
        {
            var normalizedEndpoint = GradingApiEndpoints.NormalizeEndpoint(bootstrap.GradingApiEndpoint);
            if (!GradingApiEndpoints.TryExtractSubject(normalizedEndpoint, out var subject) ||
                !GradingApiEndpoints.TryExtractProjectNumber(normalizedEndpoint, out var projectNumber))
            {
                throw new ArgumentException($"Khong the nhan dien project tu endpoint {bootstrap.GradingApiEndpoint}.");
            }

            await using var studentStream = studentFile.OpenReadStream();
            return await _xmlGradingRuleService.GradeAsync(studentStream, subject, $"project{projectNumber:00}");
        }

        private static string NormalizeMessage(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return message;
            }

            if (!message.Contains("Ã", StringComparison.Ordinal) &&
                !message.Contains("Æ", StringComparison.Ordinal) &&
                !message.Contains("â", StringComparison.Ordinal))
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
