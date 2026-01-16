using MOS.ExcelGrading.Core.DTOs;
using Microsoft.AspNetCore.Http;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IStudentService
    {
        Task<List<StudentResponse>> GetAllAsync();
        Task<StudentResponse?> GetByIdAsync(string id);
        Task<StudentResponse> CreateAsync(CreateStudentRequest request, string userId);
        Task<StudentResponse?> UpdateAsync(string id, UpdateStudentRequest request, string userId);
        Task<bool> DeleteAsync(string id);
        Task<List<StudentResponse>> GetByClassIdAsync(string classId);
        // Thêm method import
        Task<ImportStudentResult> ImportFromExcelAsync(
            IFormFile excelFile,
            string userId,
            string? teacherId = null,
            string? teacherName = null);

        Task<BulkImportResult> BulkImportAsync(
            BulkImportStudentRequest request,
            string userId);
    }
}
