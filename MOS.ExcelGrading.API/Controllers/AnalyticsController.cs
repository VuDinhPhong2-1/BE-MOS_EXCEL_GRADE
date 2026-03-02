using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
    public class AnalyticsController : ControllerBase
    {
        private readonly IAnalyticsService _analyticsService;

        public AnalyticsController(IAnalyticsService analyticsService)
        {
            _analyticsService = analyticsService;
        }

        [HttpGet("class/{classId}/overview")]
        public async Task<IActionResult> GetClassOverview(string classId)
        {
            if (!HasPermission(Permissions.ViewGrades))
                return Forbid();

            var result = await _analyticsService.GetClassOverviewAsync(classId);
            return Ok(result);
        }

        [HttpGet("class/{classId}/weak-tasks")]
        public async Task<IActionResult> GetWeakTasks(
            string classId,
            [FromQuery] string? projectEndpoint = null,
            [FromQuery] int top = 10)
        {
            if (!HasPermission(Permissions.ViewGrades))
                return Forbid();

            var result = await _analyticsService.GetWeakTasksAsync(classId, projectEndpoint, top);
            return Ok(result);
        }

        [HttpGet("class/{classId}/project-performance")]
        public async Task<IActionResult> GetProjectPerformance(string classId)
        {
            if (!HasPermission(Permissions.ViewGrades))
                return Forbid();

            var result = await _analyticsService.GetProjectPerformanceAsync(classId);
            return Ok(result);
        }

        private bool HasPermission(string permission) =>
            User.Claims.Any(c => c.Type == "permission" && c.Value == permission);
    }
}
