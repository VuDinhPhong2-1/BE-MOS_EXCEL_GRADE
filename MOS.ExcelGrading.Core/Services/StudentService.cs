// MOS.ExcelGrading.Core/Services/StudentService.cs
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Services
{
    public class StudentService : IStudentService
    {
        private readonly IMongoCollection<Student> _students;
        private readonly ILogger<StudentService> _logger;

        public StudentService(
            IMongoDatabase database,
            ILogger<StudentService> logger)
        {
            _students = database.GetCollection<Student>("students");
            _logger = logger;
        }

        public async Task<List<StudentResponse>> GetAllAsync()
        {
            try
            {
                var students = await _students
                    .Find(s => s.IsActive)
                    .SortBy(s => s.MiddleName)
                    .ThenBy(s => s.FirstName)
                    .ToListAsync();

                return students.Select(MapToResponse).ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting all students");
                throw;
            }
        }

        public async Task<StudentResponse?> GetByIdAsync(string id)
        {
            try
            {
                var student = await _students
                    .Find(s => s.Id == id && s.IsActive)
                    .FirstOrDefaultAsync();

                return student != null ? MapToResponse(student) : null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error getting student by id: {id}");
                throw;
            }
        }

        public async Task<StudentResponse> CreateAsync(CreateStudentRequest request, string userId)
        {
            try
            {
                var student = new Student
                {
                    MiddleName = request.MiddleName,
                    FirstName = request.FirstName,
                    Status = request.Status ?? "Active",
                    TeacherId = userId, // Lấy ID của user tạo student
                    ClassId = request.ClassId,
                    CreatedAt = DateTime.UtcNow,
                    CreatedBy = userId,
                    IsActive = true
                };

                await _students.InsertOneAsync(student);

                _logger.LogInformation( $"Created student: {student.Id}");

                return MapToResponse(student);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating student");
                throw;
            }
        }

        public async Task<StudentResponse?> UpdateAsync(
            string id,
            UpdateStudentRequest request,
            string userId)
        {
            try
            {
                var updateBuilder = Builders<Student>.Update
                    .Set(s => s.UpdatedAt, DateTime.UtcNow)
                    .Set(s => s.UpdatedBy, userId);

                if (!string.IsNullOrEmpty(request.MiddleName))
                    updateBuilder = updateBuilder.Set(s => s.MiddleName, request.MiddleName);

                if (!string.IsNullOrEmpty(request.FirstName))
                    updateBuilder = updateBuilder.Set(s => s.FirstName, request.FirstName);

                if (!string.IsNullOrEmpty(request.Status))
                    updateBuilder = updateBuilder.Set(s => s.Status, request.Status);

                if (!string.IsNullOrEmpty(request.ClassId))
                    updateBuilder = updateBuilder.Set(s => s.ClassId, request.ClassId);

                var result = await _students.UpdateOneAsync(
                    s => s.Id == id && s.IsActive,
                    updateBuilder
                );

                if (result.MatchedCount == 0)
                {
                    _logger.LogWarning( $"Student not found: {id}");
                    return null;
                }

                _logger.LogInformation( $"Updated student: {id}");

                var updated = await _students
                    .Find(s => s.Id == id)
                    .FirstOrDefaultAsync();

                return updated != null ? MapToResponse(updated) : null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error updating student: {id}");
                throw;
            }
        }

        public async Task<bool> DeleteAsync(string id)
        {
            try
            {
                // Soft delete
                var update = Builders<Student>.Update
                    .Set(s => s.IsActive, false)
                    .Set(s => s.UpdatedAt, DateTime.UtcNow);

                var result = await _students.UpdateOneAsync(
                    s => s.Id == id && s.IsActive,
                    update
                );

                if (result.MatchedCount == 0)
                {
                    _logger.LogWarning( $"Student not found: {id}");
                    return false;
                }

                _logger.LogInformation( $"Deleted student: {id}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error deleting student: {id}");
                throw;
            }
        }

        public async Task<ImportStudentResult> ImportFromExcelAsync(
            IFormFile excelFile,
            string userId,
            string? teacherId = null,
            string? teacherName = null)
        {
            var result = new ImportStudentResult();

            try
            {
                // Set license context for EPPlus
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using var stream = new MemoryStream();
                await excelFile.CopyToAsync(stream);

                using var package = new ExcelPackage(stream);
                var worksheet = package.Workbook.Worksheets[0]; // Sheet đầu tiên

                if (worksheet.Dimension == null)
                {
                    result.Errors.Add("File Excel trống hoặc không hợp lệ");
                    return result;
                }

                var rowCount = worksheet.Dimension.Rows;
                result.TotalRows = rowCount - 1; // Trừ header row

                // Đọc từ row 2 (bỏ qua header)
                for (int row = 2; row <= rowCount; row++)
                {
                    try
                    {
                        var middleName = worksheet.Cells[row, 1].Value?.ToString()?.Trim();
                        var firstName = worksheet.Cells[row, 2].Value?.ToString()?.Trim();
                        var classId = worksheet.Cells[row, 3].Value?.ToString()?.Trim();

                        // Validate
                        if (string.IsNullOrEmpty(middleName) || string.IsNullOrEmpty(firstName))
                        {
                            result.Errors.Add( $"Dòng {row}: Thiếu thông tin họ tên");
                            result.FailedCount++;
                            continue;
                        }

                        // Tạo student
                        var student = new Student
                        {
                            MiddleName = middleName,
                            FirstName = firstName,
                            Status = "Active",
                            TeacherId = userId, // User đang import
                            ClassId = classId,
                            CreatedAt = DateTime.UtcNow,
                            CreatedBy = userId,
                            IsActive = true
                        };

                        await _students.InsertOneAsync(student);

                        result.ImportedStudents.Add(MapToResponse(student));
                        result.SuccessCount++;

                        _logger.LogInformation( $"Imported student: {student.MiddleName} {student.FirstName}");
                    }
                    catch (Exception ex)
                    {
                        result.Errors.Add( $"Dòng {row}: {ex.Message}");
                        result.FailedCount++;
                        _logger.LogError(ex, $"Error importing row {row}");
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error importing Excel file");
                result.Errors.Add( $"Lỗi đọc file: {ex.Message}");
                return result;
            }
        }

        public async Task<BulkImportResult> BulkImportAsync(
    BulkImportStudentRequest request,
    string userId)
        {
            var result = new BulkImportResult
            {
                TotalCount = request.Students.Count
            };

            try
            {
                var studentsToInsert = new List<Student>();

                for (int i = 0; i < request.Students.Count; i++)
                {
                    var item = request.Students[i];

                    // Validate
                    if (string.IsNullOrWhiteSpace(item.MiddleName) || string.IsNullOrWhiteSpace(item.FirstName))
                    {
                        result.Errors.Add($"Dòng {i + 1}: Thiếu thông tin họ đệm hoặc tên");
                        result.FailedCount++;
                        continue;
                    }

                    var student = new Student
                    {
                        MiddleName = item.MiddleName.Trim(),
                        FirstName = item.FirstName.Trim(),
                        Status = "Active",
                        TeacherId = userId, // User đang import
                        ClassId = request.ClassId,   // DÙNG ClassId ở ngoài body!
                        CreatedAt = DateTime.UtcNow,
                        CreatedBy = userId,
                        IsActive = true
                    };

                    studentsToInsert.Add(student);
                }

                // Bulk insert để tối ưu performance
                if (studentsToInsert.Any())
                {
                    await _students.InsertManyAsync(studentsToInsert);

                    result.SuccessCount = studentsToInsert.Count;
                    result.ImportedStudents = studentsToInsert.Select(MapToResponse).ToList();

                    _logger.LogInformation($"Bulk imported {studentsToInsert.Count} students");
                }

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in BulkImportAsync");
                result.Errors.Add($"Lỗi hệ thống: {ex.Message}");
                return result;
            }
        }

        public async Task<List<StudentResponse>> GetByClassIdAsync(string classId)
        {
            try
            {
                var students = await _students
                    .Find(s => s.ClassId == classId && s.IsActive)
                    .SortBy(s => s.MiddleName)
                    .ThenBy(s => s.FirstName)
                    .ToListAsync();

                _logger.LogInformation($"Found {students.Count} students in class {classId}");

                return students.Select(MapToResponse).ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error getting students by class id: {classId}");
                throw;
            }
        }

        // Helper method to map Student to StudentResponse
        private StudentResponse MapToResponse(Student student)
        {
            return new StudentResponse
            {
                Id = student.Id ?? string.Empty,
                MiddleName = student.MiddleName,
                FirstName = student.FirstName,
                Status = student.Status,
                TeacherId = student.TeacherId,
                ClassId = student.ClassId,
                CreatedAt = student.CreatedAt,
                UpdatedAt = student.UpdatedAt,
                IsActive = student.IsActive
            };
        }
    }
}
