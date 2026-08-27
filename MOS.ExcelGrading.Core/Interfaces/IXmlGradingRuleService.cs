using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IXmlGradingRuleService
    {
        Task<List<GradingRuleSet>> GetRuleSetsAsync(string? subject = null, bool? isActive = null);
        Task<GradingRuleSet?> GetRuleSetByIdAsync(string id);
        Task<GradingRuleSet> CreateRuleSetAsync(GradingRuleSet ruleSet);
        Task<GradingRuleSet?> UpdateRuleSetAsync(string id, GradingRuleSet ruleSet);
        Task<bool> DeleteRuleSetAsync(string id);
        Task<GradingRuleSet?> AddProjectAsync(string ruleSetId, ProjectXmlRule project);
        Task<GradingRuleSet?> UpdateProjectAsync(string ruleSetId, string projectCode, ProjectXmlRule project);
        Task<GradingRuleSet?> DeleteProjectAsync(string ruleSetId, string projectCode);
        Task<GradingRuleSet?> AddTaskAsync(string ruleSetId, string projectCode, TaskXmlRule task);
        Task<GradingRuleSet?> UpdateTaskAsync(string ruleSetId, string projectCode, string taskId, TaskXmlRule task);
        Task<GradingRuleSet?> DeleteTaskAsync(string ruleSetId, string projectCode, string taskId);
        Task<GradingRuleSet?> AddConditionAsync(string ruleSetId, string projectCode, string taskId, XmlGradingCondition condition);
        Task<GradingRuleSet?> UpdateConditionAsync(string ruleSetId, string projectCode, string taskId, string conditionId, XmlGradingCondition condition);
        Task<GradingRuleSet?> DeleteConditionAsync(string ruleSetId, string projectCode, string taskId, string conditionId);
        Task<GradingRuleSet?> GetActiveRuleSetAsync(string subject, string projectCode);
        Task<XmlRuleValidationResult> ValidateRuleSetAsync(GradingRuleSet ruleSet);
        Task<GradingRuleSet> SeedProject22Task1RuleSetAsync();
        Task<GradingResult> GradeAsync(Stream studentFile, string subject, string projectCode);
    }
}