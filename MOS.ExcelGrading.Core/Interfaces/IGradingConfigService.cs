using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IGradingConfigService
    {
        Task<List<GradingConfigListItemDto>> GetAllAsync(string? subject = null, string? status = null);
        Task<GradingConfigDetailDto?> GetByIdAsync(string id);
        Task<GradingConfigDetailDto?> GetActiveByEndpointAsync(string gradingApiEndpoint);
        Task<GradingConfigDetailDto> ImportFromCodeAsync(ImportGradingConfigFromCodeRequest request, string userId);
        Task<GradingConfigDetailDto> UpdateAsync(string id, UpdateGradingConfigRequest request, string userId);
        Task<GradingConfigDetailDto> PublishAsync(string id, PublishGradingConfigRequest request, string userId);
        Task<List<GradingConfigVersionDto>> GetVersionsAsync(string id);
        Task<GradingConfigDetailDto?> GetVersionSnapshotAsync(string id, int version);
        Task<GradingConfigDetailDto> RestoreVersionAsync(string id, int version, RestoreGradingConfigVersionRequest request, string userId);
        Task<List<GradingConfigTestRunDto>> GetTestRunsAsync(string id);
        Task<GradingConfigTestRunDto> CreateTestRunAsync(string id, GradingResult result, string fileName, bool usedOverride, string userId, string? error = null);
        Task<List<GradingRuleTypeDto>> GetRuleTypesAsync();
    }
}