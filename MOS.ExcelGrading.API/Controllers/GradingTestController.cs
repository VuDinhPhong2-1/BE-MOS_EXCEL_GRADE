using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Security.Claims;
using System.Text.RegularExpressions;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/grading-test")]
    [Authorize(Roles = $"{UserRoles.Admin},{UserRoles.Teacher}")]
    public class GradingTestController : ControllerBase
    {
        private static readonly Regex ProjectCodeRegex = new(
            @"^project(?<number>\d{1,2})(?:-(excel|word))?$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private const int MaxSupportedProject = 22;
        private static readonly HashSet<int> SupportedProjects = Enumerable.Range(1, 16)
            .Append(18)
            .Append(20)
            .Append(22)
            .ToHashSet();

        private readonly IGradingService _gradingService;
        private readonly IGradingTestBugNoteService _gradingTestBugNoteService;
        private readonly ILogger<GradingTestController> _logger;

        public GradingTestController(
            IGradingService gradingService,
            IGradingTestBugNoteService gradingTestBugNoteService,
            ILogger<GradingTestController> logger)
        {
            _gradingService = gradingService;
            _gradingTestBugNoteService = gradingTestBugNoteService;
            _logger = logger;
        }

        [HttpGet("projects")]
        public IActionResult GetSupportedProjects()
        {
            var excelProjects = SupportedProjects
                .OrderBy(projectNumber => projectNumber)
                .Select(i => new
                {
                    code = $"project{i:00}-excel",
                    endpoint = GradingApiEndpoints.ToExcelProjectEndpoint(i),
                    displayName = $"Project {i:00} - Excel",
                    fileType = "excel"
                })
                .ToList();

            // Add Word projects (use /api/grading/word/ endpoint)
            var wordProjects = SupportedProjects
                .OrderBy(projectNumber => projectNumber)
                .Select(i => new
                {
                    code = $"project{i:00}-word",
                    endpoint = GradingApiEndpoints.ToWordProjectEndpoint(i),
                    displayName = $"Project {i:00} - Word",
                    fileType = "word"
                })
                .ToList();

            var allProjects = excelProjects.Concat(wordProjects).ToList();
            return Ok(allProjects);
        }

        [HttpGet("bug-notes")]
        public async Task<IActionResult> GetBugNotes([FromQuery] string? projectCode = null)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrWhiteSpace(userId))
                {
                    return Unauthorized();
                }

                var notes = await _gradingTestBugNoteService.GetByUserAsync(userId, projectCode);
                return Ok(notes);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error loading grading test bug notes");
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        [HttpPost("bug-notes")]
        public async Task<IActionResult> CreateBugNote([FromBody] CreateGradingTestBugNoteRequest request)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrWhiteSpace(userId))
                {
                    return Unauthorized();
                }

                var createdNote = await _gradingTestBugNoteService.CreateAsync(request, userId);
                return Ok(createdNote);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating grading test bug note");
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
        }

        [HttpDelete("bug-notes/{id}")]
        public async Task<IActionResult> DeleteBugNote(string id)
        {
            try
            {
                var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrWhiteSpace(userId))
                {
                    return Unauthorized();
                }

                var deleted = await _gradingTestBugNoteService.DeleteAsync(id, userId);
                if (!deleted)
                {
                    return NotFound(new { message = "Không tìm thấy bug note." });
                }

                return NoContent();
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting grading test bug note {Id}", id);
                return StatusCode(500, "Lỗi máy chủ nội bộ");
            }
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
                        error = $"Project không hợp lệ. Chỉ hỗ trợ project01 đến project{MaxSupportedProject:00}, trừ project17, project19 và project21."
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

        [HttpPost("word/{projectCode}")]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GradeWordProject(
            string projectCode,
            [FromForm] IFormFile studentFile)
        {
            try
            {
                if (studentFile == null)
                {
                    return BadRequest(new { error = "Can cung cap file: studentFile" });
                }

                if (!TryExtractSupportedWordProjectNumber(projectCode, out var projectNumber))
                {
                    return BadRequest(new
                    {
                        error = $"Project Word khong hop le. Chi ho tro project01 den project{MaxSupportedProject:00}, tru project17, project19 va project21."
                    });
                }

                if (!IsAcceptedWordInputForProject(studentFile, projectNumber))
                {
                    if (projectNumber == 7)
                    {
                        return BadRequest(new
                        {
                            error = "Word Project 07 yeu cau .docx hoac file plain text ten Memo.txt."
                        });
                    }

                    return BadRequest(new { error = "File phai co dinh dang .docx (Word OpenXML)." });
                }

                using var studentStream = studentFile.OpenReadStream();
                var result = await _gradingService.GradeWordProjectAsync(projectNumber, studentStream, studentFile.FileName);
                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error grading test word project {ProjectCode}", projectCode);
                return StatusCode(500, new { error = "Loi he thong khi cham diem word test" });
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
                12 => await _gradingService.GradeProject12Async(studentStream),
                13 => await _gradingService.GradeProject13Async(studentStream),
                14 => await _gradingService.GradeProject14Async(studentStream),
                15 => await _gradingService.GradeProject15Async(studentStream),
                16 => await _gradingService.GradeProject16Async(studentStream),
                18 => await _gradingService.GradeProject18Async(studentStream),
                20 => await _gradingService.GradeProject20Async(studentStream),
                22 => await _gradingService.GradeProject22Async(studentStream),
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

            if (!SupportedProjects.Contains(value))
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

        private static bool TryExtractSupportedWordProjectNumber(string projectCode, out int projectNumber)
        {
            projectNumber = 0;
            var normalizedEndpoint = GradingApiEndpoints.NormalizeEndpoint($"{GradingApiSubjects.Word}/{projectCode}");
            if (!normalizedEndpoint.StartsWith($"{GradingApiSubjects.Word}/project", StringComparison.Ordinal))
            {
                return false;
            }

            if (!GradingApiEndpoints.IsValidEndpoint(normalizedEndpoint))
            {
                return false;
            }

            if (!GradingApiEndpoints.TryExtractProjectNumber(normalizedEndpoint, out var parsedProjectNumber))
            {
                return false;
            }

            if (!SupportedProjects.Contains(parsedProjectNumber))
            {
                return false;
            }

            projectNumber = parsedProjectNumber;
            return true;
        }

        private static bool IsAcceptedWordInputForProject(IFormFile file, int projectNumber)
        {
            if (projectNumber == 7 && IsMemoPlainTextFile(file))
            {
                return true;
            }

            return IsWordOpenXmlFile(file);
        }

        private static bool IsWordOpenXmlFile(IFormFile file)
        {
            var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
            return extension == ".docx";
        }

        private static bool IsMemoPlainTextFile(IFormFile file)
        {
            var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (extension != ".txt")
            {
                return false;
            }

            var baseName = Path.GetFileNameWithoutExtension(file.FileName) ?? string.Empty;
            return string.Equals(baseName, "memo", StringComparison.OrdinalIgnoreCase);
        }
    }
}
