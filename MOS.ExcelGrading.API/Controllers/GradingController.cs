using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.Interfaces;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class GradingController : ControllerBase
    {
        private readonly IGradingService _gradingService;
        private readonly ILogger<GradingController> _logger;

        public GradingController(
            IGradingService gradingService,
            ILogger<GradingController> logger)
        {
            _gradingService = gradingService;
            _logger = logger;
        }

        [HttpPost("project09")]
        [RequestSizeLimit(524288000)] // ✅ 500MB
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)] // ✅ 500MB
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeProject09(
            [FromForm] IFormFile studentFile,
            [FromForm] IFormFile answerFile)
        {
            try
            {
                if (studentFile == null || answerFile == null)
                {
                    return BadRequest(new { error = "Cần cung cấp cả 2 file: studentFile và answerFile" });
                }

                // Log kích thước file
                _logger.LogInformation($"Student file: {studentFile.FileName} ({studentFile.Length:N0} bytes)");
                _logger.LogInformation($"Answer file: {answerFile.FileName} ({answerFile.Length:N0} bytes)");

                if (!IsExcelFile(studentFile) || !IsExcelFile(answerFile))
                {
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });
                }

                using var studentStream = studentFile.OpenReadStream();
                using var answerStream = answerFile.OpenReadStream();

                var result = await _gradingService.GradeProject09Async(studentStream, answerStream);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi chấm điểm Project 09");
                return StatusCode(500, new { error = $"Lỗi hệ thống: {ex.Message}" });
            }
        }

        [HttpGet("health")]
        public IActionResult HealthCheck()
        {
            return Ok(new
            {
                status = "OK",
                timestamp = DateTime.UtcNow,
                maxUploadSize = "500MB"
            });
        }

        private bool IsExcelFile(IFormFile file)
        {
            var extension = Path.GetExtension(file.FileName).ToLower();
            return extension == ".xlsx" || extension == ".xlsm";
        }
    }
}
