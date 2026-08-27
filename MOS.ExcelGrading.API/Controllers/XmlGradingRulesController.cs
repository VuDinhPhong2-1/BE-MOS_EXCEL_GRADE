using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/admin/xml-grading-rules")]
    [Authorize(Roles = UserRoles.Admin)]
    public class XmlGradingRulesController : ControllerBase
    {
        private readonly IXmlGradingRuleService _xmlGradingRuleService;
        private readonly ILogger<XmlGradingRulesController> _logger;

        public XmlGradingRulesController(
            IXmlGradingRuleService xmlGradingRuleService,
            ILogger<XmlGradingRulesController> logger)
        {
            _xmlGradingRuleService = xmlGradingRuleService;
            _logger = logger;
        }

        [HttpGet]
        public async Task<ActionResult<List<GradingRuleSet>>> GetRuleSets([FromQuery] string? subject, [FromQuery] bool? isActive)
        {
            var ruleSets = await _xmlGradingRuleService.GetRuleSetsAsync(subject, isActive);
            return Ok(ruleSets);
        }

        [HttpGet("{id}")]
        public async Task<ActionResult<GradingRuleSet>> GetRuleSetById(string id)
        {
            var ruleSet = await _xmlGradingRuleService.GetRuleSetByIdAsync(id);
            if (ruleSet == null)
            {
                return NotFound(new { message = "Không tìm thấy XML grading ruleset." });
            }

            return Ok(ruleSet);
        }

        [HttpPost]
        public async Task<ActionResult<GradingRuleSet>> CreateRuleSet([FromBody] GradingRuleSet ruleSet)
        {
            try
            {
                var created = await _xmlGradingRuleService.CreateRuleSetAsync(ruleSet);
                return CreatedAtAction(nameof(GetRuleSetById), new { id = created.Id }, created);
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        [HttpPut("{id}")]
        public async Task<ActionResult<GradingRuleSet>> UpdateRuleSet(string id, [FromBody] GradingRuleSet ruleSet)
        {
            try
            {
                var updated = await _xmlGradingRuleService.UpdateRuleSetAsync(id, ruleSet);
                if (updated == null)
                {
                    return NotFound(new { message = "Không tìm thấy XML grading ruleset." });
                }

                return Ok(updated);
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteRuleSet(string id)
        {
            var deleted = await _xmlGradingRuleService.DeleteRuleSetAsync(id);
            if (!deleted)
            {
                return NotFound(new { message = "Không tìm thấy XML grading ruleset." });
            }

            return NoContent();
        }

        [HttpPost("{ruleSetId}/projects")]
        public async Task<ActionResult<GradingRuleSet>> AddProject(string ruleSetId, [FromBody] ProjectXmlRule project)
        {
            return await ExecuteNestedMutation(() => _xmlGradingRuleService.AddProjectAsync(ruleSetId, project));
        }

        [HttpPut("{ruleSetId}/projects/{projectCode}")]
        public async Task<ActionResult<GradingRuleSet>> UpdateProject(string ruleSetId, string projectCode, [FromBody] ProjectXmlRule project)
        {
            return await ExecuteNestedMutation(() => _xmlGradingRuleService.UpdateProjectAsync(ruleSetId, projectCode, project));
        }

        [HttpDelete("{ruleSetId}/projects/{projectCode}")]
        public async Task<ActionResult<GradingRuleSet>> DeleteProject(string ruleSetId, string projectCode)
        {
            return await ExecuteNestedMutation(() => _xmlGradingRuleService.DeleteProjectAsync(ruleSetId, projectCode));
        }

        [HttpPost("{ruleSetId}/projects/{projectCode}/tasks")]
        public async Task<ActionResult<GradingRuleSet>> AddTask(string ruleSetId, string projectCode, [FromBody] TaskXmlRule task)
        {
            return await ExecuteNestedMutation(() => _xmlGradingRuleService.AddTaskAsync(ruleSetId, projectCode, task));
        }

        [HttpPut("{ruleSetId}/projects/{projectCode}/tasks/{taskId}")]
        public async Task<ActionResult<GradingRuleSet>> UpdateTask(string ruleSetId, string projectCode, string taskId, [FromBody] TaskXmlRule task)
        {
            return await ExecuteNestedMutation(() => _xmlGradingRuleService.UpdateTaskAsync(ruleSetId, projectCode, taskId, task));
        }

        [HttpDelete("{ruleSetId}/projects/{projectCode}/tasks/{taskId}")]
        public async Task<ActionResult<GradingRuleSet>> DeleteTask(string ruleSetId, string projectCode, string taskId)
        {
            return await ExecuteNestedMutation(() => _xmlGradingRuleService.DeleteTaskAsync(ruleSetId, projectCode, taskId));
        }

        [HttpPost("{ruleSetId}/projects/{projectCode}/tasks/{taskId}/conditions")]
        public async Task<ActionResult<GradingRuleSet>> AddCondition(string ruleSetId, string projectCode, string taskId, [FromBody] XmlGradingCondition condition)
        {
            return await ExecuteNestedMutation(() => _xmlGradingRuleService.AddConditionAsync(ruleSetId, projectCode, taskId, condition));
        }

        [HttpPut("{ruleSetId}/projects/{projectCode}/tasks/{taskId}/conditions/{conditionId}")]
        public async Task<ActionResult<GradingRuleSet>> UpdateCondition(string ruleSetId, string projectCode, string taskId, string conditionId, [FromBody] XmlGradingCondition condition)
        {
            return await ExecuteNestedMutation(() => _xmlGradingRuleService.UpdateConditionAsync(ruleSetId, projectCode, taskId, conditionId, condition));
        }

        [HttpDelete("{ruleSetId}/projects/{projectCode}/tasks/{taskId}/conditions/{conditionId}")]
        public async Task<ActionResult<GradingRuleSet>> DeleteCondition(string ruleSetId, string projectCode, string taskId, string conditionId)
        {
            return await ExecuteNestedMutation(() => _xmlGradingRuleService.DeleteConditionAsync(ruleSetId, projectCode, taskId, conditionId));
        }

        [HttpGet("{subject}/{projectCode}/active")]
        public async Task<ActionResult<GradingRuleSet>> GetActiveRuleSet(string subject, string projectCode)
        {
            var ruleSet = await _xmlGradingRuleService.GetActiveRuleSetAsync(subject, projectCode);
            if (ruleSet == null)
            {
                return NotFound(new { message = $"Không tìm thấy XML grading ruleset active cho {subject}/{projectCode}." });
            }

            return Ok(ruleSet);
        }

        [HttpPost("validate")]
        public async Task<ActionResult<XmlRuleValidationResult>> ValidateRuleSet([FromBody] GradingRuleSet ruleSet)
        {
            var validation = await _xmlGradingRuleService.ValidateRuleSetAsync(ruleSet);
            return validation.IsValid ? Ok(validation) : BadRequest(validation);
        }

        [HttpPost("seed/excel/project22/task1")]
        public async Task<ActionResult<GradingRuleSet>> SeedProject22Task1()
        {
            try
            {
                var ruleSet = await _xmlGradingRuleService.SeedProject22Task1RuleSetAsync();
                return Ok(ruleSet);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to seed XML grading rules for Excel project22 task1.");
                return BadRequest(new { message = ex.Message });
            }
        }

        [HttpPost("grade/{subject}/{projectCode}")]
        [RequestSizeLimit(104_857_600)]
        public async Task<ActionResult<GradingResult>> GradeWithXmlRules(string subject, string projectCode, IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest(new { message = "Vui lòng upload file Office cần chấm." });
            }

            try
            {
                await using var stream = file.OpenReadStream();
                var result = await _xmlGradingRuleService.GradeAsync(stream, subject, projectCode);
                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to grade file with XML rules for {Subject}/{ProjectCode}.", subject, projectCode);
                return BadRequest(new { message = ex.Message });
            }
        }

        private async Task<ActionResult<GradingRuleSet>> ExecuteNestedMutation(Func<Task<GradingRuleSet?>> mutation)
        {
            try
            {
                var ruleSet = await mutation();
                if (ruleSet == null)
                {
                    return NotFound(new { message = "Không tìm thấy XML grading ruleset hoặc nested resource." });
                }

                return Ok(ruleSet);
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }
    }
}