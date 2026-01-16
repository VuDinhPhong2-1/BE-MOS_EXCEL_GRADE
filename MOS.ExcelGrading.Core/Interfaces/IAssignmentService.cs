// MOS.ExcelGrading.Core/Interfaces/IAssignmentService.cs
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IAssignmentService
    {
        Task<List<Assignment>> GetAssignmentsByClassIdAsync(string classId);
        Task<List<AssignmentWithStatsResponse>> GetAssignmentsWithStatsByClassIdAsync(string classId);
        Task<Assignment?> GetAssignmentByIdAsync(string id);
        Task<Assignment> CreateAssignmentAsync(CreateAssignmentRequest request, string userId);
        Task<Assignment?> UpdateAssignmentAsync(string id, UpdateAssignmentRequest request, string userId);
        Task<bool> DeleteAssignmentAsync(string id, string userId);
        Task<bool> CanUserAccessAssignment(string assignmentId, string userId);
    }
}
