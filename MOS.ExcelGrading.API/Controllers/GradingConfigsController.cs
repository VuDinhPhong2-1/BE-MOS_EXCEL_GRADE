using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Security.Claims;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/admin/grading-configs")]
    [Authorize(Roles = UserRoles.Admin)]
    public class GradingConfigsController : ControllerBase
    {
        private readonly IGradingConfigService _gradingConfigService;
        private readonly IGradingService _gradingService;
        private readonly ILogger<GradingConfigsController> _logger;

        public GradingConfigsController(
            IGradingConfigService gradingConfigService,
            IGradingService gradingService,
            ILogger<GradingConfigsController> logger)
        {
            _gradingConfigService = gradingConfigService;
            _gradingService = gradingService;
            _logger = logger;
        }

        [HttpGet]
        public async Task<IActionResult> GetAll([FromQuery] string? subject = null, [FromQuery] string? status = null)
        {
            try
            {
                var configs = await _gradingConfigService.GetAllAsync(subject, status);
                return Ok(configs);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        [HttpGet("rule-types")]
        public async Task<IActionResult> GetRuleTypes()
        {
            var ruleTypes = await _gradingConfigService.GetRuleTypesAsync();
            return Ok(ruleTypes);
        }

        [HttpGet("active")]
        public async Task<IActionResult> GetActiveByEndpoint([FromQuery] string gradingApiEndpoint)
        {
            try
            {
                var config = await _gradingConfigService.GetActiveByEndpointAsync(gradingApiEndpoint);
                if (config == null)
                {
                    return NotFound(new { message = "Không tìm thấy grading config active." });
                }

                return Ok(config);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        [HttpGet("{id}")]
        public async Task<IActionResult> GetById(string id)
        {
            try
            {
                var config = await _gradingConfigService.GetByIdAsync(id);
                if (config == null)
                {
                    return NotFound(new { message = "Không tìm thấy grading config." });
                }

                return Ok(config);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        [HttpPost("import-from-code")]
        public async Task<IActionResult> ImportFromCode([FromBody] ImportGradingConfigFromCodeRequest request)
        {
            try
            {
                var userId = GetCurrentUserId();
                var config = await _gradingConfigService.ImportFromCodeAsync(request, userId);
                return Ok(config);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error importing grading config from code");
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> Update(string id, [FromBody] UpdateGradingConfigRequest request)
        {
            try
            {
                var userId = GetCurrentUserId();
                var config = await _gradingConfigService.UpdateAsync(id, request, userId);
                return Ok(config);
            }
            catch (KeyNotFoundException ex)
            {
                return NotFound(new { message = ex.Message });
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        [HttpPost("{id}/publish")]
        public async Task<IActionResult> Publish(string id, [FromBody] PublishGradingConfigRequest request)
        {
            try
            {
                var userId = GetCurrentUserId();
                var config = await _gradingConfigService.PublishAsync(id, request, userId);
                return Ok(config);
            }
            catch (KeyNotFoundException ex)
            {
                return NotFound(new { message = ex.Message });
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        [HttpGet("{id}/test-runs")]
        public async Task<IActionResult> GetTestRuns(string id)
        {
            try
            {
                var testRuns = await _gradingConfigService.GetTestRunsAsync(id);
                return Ok(testRuns);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        [HttpPost("{id}/test")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> TestConfig(
            string id,
            [FromForm] IFormFile studentFile)
        {
            try
            {
                if (studentFile == null)
                {
                    return BadRequest(new { error = "Cần cung cấp file: studentFile" });
                }

                var config = await _gradingConfigService.GetByIdAsync(id);
                if (config == null)
                {
                    return NotFound(new { message = "Không tìm thấy grading config." });
                }

                if (!GradingApiEndpoints.TryExtractSubject(config.GradingApiEndpoint, out var subject) ||
                    !GradingApiEndpoints.TryExtractProjectNumber(config.GradingApiEndpoint, out var projectNumber))
                {
                    return BadRequest(new { message = "Grading config có endpoint không hợp lệ." });
                }

                GradingResult result;
                using var studentStream = studentFile.OpenReadStream();

                if (string.Equals(subject, AssignmentFileSubjects.Word, StringComparison.OrdinalIgnoreCase))
                {
                    result = await _gradingService.GradeWordProjectAsync(projectNumber, studentStream, studentFile.FileName);
                }
                else if (string.Equals(subject, AssignmentFileSubjects.Excel, StringComparison.OrdinalIgnoreCase))
                {
                    result = await GradeExcelByProjectNumberAsync(projectNumber, studentStream);
                }
                else
                {
                    return BadRequest(new { message = "Subject grading config không được hỗ trợ." });
                }

                await _gradingConfigService.CreateTestRunAsync(id, result, studentFile.FileName, false, GetCurrentUserId());
                return Ok(result);
            }
            catch (KeyNotFoundException ex)
            {
                return NotFound(new { message = ex.Message });
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error testing grading config {Id}", id);
                return StatusCode(500, new { error = "Lỗi hệ thống khi test grading config" });
            }
        }

        [HttpGet("{id}/versions")]
        public async Task<IActionResult> GetVersions(string id)
        {
            try
            {
                var versions = await _gradingConfigService.GetVersionsAsync(id);
                return Ok(versions);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        [HttpGet("{id}/versions/{version:int}")]
        public async Task<IActionResult> GetVersionSnapshot(string id, int version)
        {
            try
            {
                var snapshot = await _gradingConfigService.GetVersionSnapshotAsync(id, version);
                if (snapshot == null)
                {
                    return NotFound(new { message = "Không tìm thấy version grading config." });
                }

                return Ok(snapshot);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        [HttpPost("{id}/versions/{version:int}/restore")]
        public async Task<IActionResult> RestoreVersion(string id, int version, [FromBody] RestoreGradingConfigVersionRequest request)
        {
            try
            {
                var userId = GetCurrentUserId();
                var restored = await _gradingConfigService.RestoreVersionAsync(id, version, request, userId);
                return Ok(restored);
            }
            catch (KeyNotFoundException ex)
            {
                return NotFound(new { message = ex.Message });
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        private async Task<GradingResult> GradeExcelByProjectNumberAsync(int projectNumber, Stream studentStream)
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
                12 => await _gradingService.GradeProject12Async(studentStream),
                13 => await _gradingService.GradeProject13Async(studentStream),
                14 => await _gradingService.GradeProject14Async(studentStream),
                15 => await _gradingService.GradeProject15Async(studentStream),
                16 => await _gradingService.GradeProject16Async(studentStream),
                18 => await _gradingService.GradeProject18Async(studentStream),
                20 => await _gradingService.GradeProject20Async(studentStream),
                22 => await _gradingService.GradeProject22Async(studentStream),
                _ => throw new InvalidOperationException($"Chưa hỗ trợ chấm Excel project {projectNumber:00}.")
            };
        }

        private string GetCurrentUserId()
        {
            var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (string.IsNullOrWhiteSpace(userId))
            {
                throw new InvalidOperationException("Không xác định được người dùng hiện tại.");
            }

            return userId;
        }
    }
}