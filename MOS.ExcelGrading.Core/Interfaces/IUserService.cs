using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IUserService
    {
        Task<User?> RegisterAsync(string email, string username, string password, string? role = null, string? fullName = null);
        Task<AuthResponse?> LoginAsync(string username, string password);
        Task<AuthResponse?> RefreshTokenAsync(string refreshToken);
        Task<User?> GetUserByIdAsync(string id);
        Task<User?> GetUserByUsernameAsync(string username);
        Task<bool> RevokeRefreshTokenAsync(string userId);
    }
}
