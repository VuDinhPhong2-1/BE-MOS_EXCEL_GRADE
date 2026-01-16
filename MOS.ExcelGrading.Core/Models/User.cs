// Core/Models/User.cs
using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.Models
{
    public class User
    {
        // ========== QUAN HỆ VỚI SCHOOL ==========
        [BsonRepresentation(BsonType.ObjectId)]
        public List<string> SchoolIds { get; set; } = new List<string>();

        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string? Id { get; set; }

        [Required(ErrorMessage = "Email là bắt buộc")]
        [EmailAddress(ErrorMessage = "Email không hợp lệ")]
        public string Email { get; set; } = string.Empty;

        [Required(ErrorMessage = "Username là bắt buộc")]
        [StringLength(50, MinimumLength = 3)]
        public string Username { get; set; } = string.Empty;

        [Required(ErrorMessage = "Password là bắt buộc")]
        public string PasswordHash { get; set; } = string.Empty;

        // ========== PHÂN QUYỀN ==========
        [Required]
        public string Role { get; set; } = UserRoles.Teacher;
        public List<string> Permissions { get; set; } = new List<string>();

        // ========== THÔNG TIN BỔ SUNG ==========
        public string? FullName { get; set; }
        public string? PhoneNumber { get; set; }
        public string? Avatar { get; set; }

        // ========== METADATA ==========
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime? LastLogin { get; set; }
        public bool IsActive { get; set; } = true;
        public bool IsEmailVerified { get; set; } = false;

        // ========== REFRESH TOKEN ==========
        public string? RefreshToken { get; set; }
        public DateTime? RefreshTokenExpiry { get; set; }
    }

    // ========== ĐỊNH NGHĨA ROLES ==========
    public static class UserRoles
    {
        public const string Admin = "Admin";
        public const string Teacher = "Teacher";
        public const string Student = "Student";

        public static List<string> GetAllRoles() => new List<string> { Admin, Teacher, Student };
        public static bool IsValidRole(string role) => GetAllRoles().Contains(role);
    }

    // ========== ĐỊNH NGHĨA PERMISSIONS ==========
    public static class Permissions
    {
        // User Management
        public const string ViewUsers = "users.view";
        public const string CreateUsers = "users.create";
        public const string EditUsers = "users.edit";
        public const string DeleteUsers = "users.delete";

        // Grading Management
        public const string ViewGrades = "grades.view";
        public const string CreateGrades = "grades.create";
        public const string EditGrades = "grades.edit";
        public const string DeleteGrades = "grades.delete";
        public const string ExportGrades = "grades.export";

        // Project Management
        public const string ViewProjects = "projects.view";
        public const string CreateProjects = "projects.create";
        public const string EditProjects = "projects.edit";
        public const string DeleteProjects = "projects.delete";

        // School Management
        public const string ViewSchools = "schools.view";
        public const string CreateSchools = "schools.create";
        public const string EditSchools = "schools.edit";
        public const string DeleteSchools = "schools.delete";

        // System
        public const string ViewSystemLogs = "system.logs.view";
        public const string ManageSettings = "system.settings.manage";

        // Student Management
        public const string ViewStudents = "students.view";
        public const string CreateStudents = "students.create";
        public const string EditStudents = "students.edit";
        public const string DeleteStudents = "students.delete";
        public const string ImportStudents = "students.import";
        public const string BulkImportStudents = "students.bulkimport";

        public static Dictionary<string, List<string>> GetRolePermissions() => new()
        {
            [UserRoles.Admin] = new List<string>
            {
                ViewUsers, CreateUsers, EditUsers, DeleteUsers,
                ViewGrades, CreateGrades, EditGrades, DeleteGrades, ExportGrades,
                ViewProjects, CreateProjects, EditProjects, DeleteProjects,
                ViewSchools, CreateSchools, EditSchools, DeleteSchools,
                ViewSystemLogs, ManageSettings
            },
            [UserRoles.Teacher] = new List<string>
            {
                ViewUsers,
                ViewGrades, CreateGrades, EditGrades, ExportGrades,
                ViewProjects, CreateProjects, EditProjects,
                ViewSchools, CreateSchools, EditSchools
            },
            [UserRoles.Student] = new List<string>
            {
                ViewGrades, ViewProjects, ViewSchools
            }
        };
    }
}
