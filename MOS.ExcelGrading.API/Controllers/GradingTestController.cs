using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Text.RegularExpressions;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/grading-test")]
    [Authorize] // Chỉ cần đăng nhập, không yêu cầu permission chấm điểm.
    public class GradingTestController : ControllerBase
    {
        private static readonly Regex ProjectCodeRegex = new(
            @"^project(?<number>\d{1,2})$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private const int MaxSupportedProject = 11;

        private readonly IGradingService _gradingService;
        private readonly ILogger<GradingTestController> _logger;

        public GradingTestController(
            IGradingService gradingService,
            ILogger<GradingTestController> logger)
        {
            _gradingService = gradingService;
            _logger = logger;
        }

        [HttpGet("projects")]
        public IActionResult GetSupportedProjects()
        {
            var projects = Enumerable.Range(1, MaxSupportedProject)
                .Select(i => new
                {
                    code = $"project{i:00}",
                    endpoint = GradingApiEndpoints.ToExcelProjectEndpoint(i),
                    displayName = $"Project {i:00} - Excel"
                })
                .ToList();

            return Ok(projects);
        }

        [HttpPost("excel/{projectCode}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeExcelProject(
            string projectCode,
            [FromForm] IFormFile studentFile)
        {
            try
            {
                if (studentFile == null)
                {
                    return BadRequest(new { error = "Cần cung cấp file: studentFile" });
                }

                if (!IsExcelFile(studentFile))
                {
                    return BadRequest(new { error = "File phải có định dạng .xlsx hoặc .xlsm" });
                }

                if (!TryExtractProjectNumber(projectCode, out var projectNumber))
                {
                    return BadRequest(new
                    {
                        error = $"Project không hợp lệ. Chỉ hỗ trợ project01 đến project{MaxSupportedProject:00}."
                    });
                }

                using var studentStream = studentFile.OpenReadStream();
                var result = await GradeByProjectNumberAsync(projectNumber, studentStream);
                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading test project {ProjectCode}", projectCode);
                return StatusCode(500, new { error = "Lỗi hệ thống khi chấm điểm test" });
            }
        }

        private async Task<GradingResult> GradeByProjectNumberAsync(int projectNumber, Stream studentStream)
        {
            return projectNumber switch
            {
                1 => await _gradingService.GradeProject01Async(studentStream),
                2 => await _gradingService.GradeProject02Async(studentStream),
                3 => await _gradingService.GradeProject03Async(studentStream),
                4 => await _gradingService.GradeProject04Async(studentStream),
                5 => await _gradingService.GradeProject05Async(studentStream),
                6 => await _gradingService.GradeProject06Async(studentStream),
                7 => await _gradingService.GradeProject07Async(studentStream),
                8 => await _gradingService.GradeProject08Async(studentStream),
                9 => await _gradingService.GradeProject09Async(studentStream),
                10 => await _gradingService.GradeProject10Async(studentStream),
                11 => await _gradingService.GradeProject11Async(studentStream),
                _ => throw new ArgumentOutOfRangeException(nameof(projectNumber))
            };
        }

        private static bool TryExtractProjectNumber(string projectCode, out int projectNumber)
        {
            projectNumber = 0;
            var normalized = (projectCode ?? string.Empty)
                .Trim()
                .Replace("\\", "/", StringComparison.Ordinal)
                .Trim('/')
                .ToLowerInvariant();

            if (normalized.StartsWith("excel/", StringComparison.Ordinal))
            {
                normalized = normalized["excel/".Length..];
            }

            var match = ProjectCodeRegex.Match(normalized);
            if (!match.Success)
            {
                return false;
            }

            if (!int.TryParse(match.Groups["number"].Value, out var value))
            {
                return false;
            }

            if (value < 1 || value > MaxSupportedProject)
            {
                return false;
            }

            projectNumber = value;
            return true;
        }

        private static bool IsExcelFile(IFormFile file)
        {
            var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
            return extension == ".xlsx" || extension == ".xlsm";
        }
    }
}
