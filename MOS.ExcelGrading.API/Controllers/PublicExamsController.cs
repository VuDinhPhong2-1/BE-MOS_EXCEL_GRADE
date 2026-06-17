using System.Text;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.Interfaces;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/public/exams")]
    public class PublicExamsController : ControllerBase
    {
        private static readonly Encoding Latin1 = Encoding.GetEncoding(1252);

        private readonly IExamPublicationService _examPublicationService;
        private readonly ILogger<PublicExamsController> _logger;

        public PublicExamsController(
            IExamPublicationService examPublicationService,
            ILogger<PublicExamsController> logger)
        {
            _examPublicationService = examPublicationService;
            _logger = logger;
        }

        [HttpGet("{token}")]
        public async Task<IActionResult> GetByToken(string token)
        {
            try
            {
                var publication = await _examPublicationService.GetPublicExamPublicationByTokenAsync(token);
                if (publication == null)
                {
                    return NotFound(new { message = "Không tìm thấy ca thi" });
                }

                return Ok(publication);
            }
            catch (ArgumentException ex)
            {
                var message = NormalizeMessage(ex.Message);
                _logger.LogWarning("Validation error getting public exam publication: {Message}", message);
                return BadRequest(new { message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting public exam publication");
                return StatusCode(500, new { message = "Lỗi máy chủ nội bộ" });
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
