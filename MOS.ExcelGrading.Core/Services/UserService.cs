// Core/Services/UserService.cs
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Tokens;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Extensions.Configuration;

namespace MOS.ExcelGrading.Core.Services
{
    public class UserService : IUserService
    {
        private readonly IMongoCollection<User> _users;
        private readonly string _jwtSecretKey;
        private readonly string _jwtIssuer;
        private readonly string _jwtAudience;
        private readonly int _jwtExpiryMinutes;
        private readonly int _refreshTokenExpiryDays;

        public UserService(
            IOptions<MongoDbSettings> mongoSettings,
            IConfiguration configuration)
        {
            var client = new MongoClient(mongoSettings.Value.ConnectionString);
            var database = client.GetDatabase(mongoSettings.Value.DatabaseName);
            _users = database.GetCollection<User>(mongoSettings.Value.UsersCollectionName);

            // Tạo index unique cho Username và Email
            var indexKeysDefinition = Builders<User>.IndexKeys.Ascending(u => u.Username);
            var indexOptions = new CreateIndexOptions { Unique = true };
            var indexModel = new CreateIndexModel<User>(indexKeysDefinition, indexOptions);
            _users.Indexes.CreateOne(indexModel);

            var emailIndexKeys = Builders<User>.IndexKeys.Ascending(u => u.Email);
            var emailIndexModel = new CreateIndexModel<User>(emailIndexKeys, new CreateIndexOptions { Unique = true });
            _users.Indexes.CreateOne(emailIndexModel);

            // JWT Configuration
            _jwtSecretKey = configuration["JwtSettings:SecretKey"] ?? throw new ArgumentNullException("JWT SecretKey");
            _jwtIssuer = configuration["JwtSettings:Issuer"] ?? "MOS.ExcelGrading";
            _jwtAudience = configuration["JwtSettings:Audience"] ?? "MOS.ExcelGrading.Users";
            _jwtExpiryMinutes = int.Parse(configuration["JwtSettings:ExpiryMinutes"] ?? "60");
            _refreshTokenExpiryDays = int.Parse(configuration["JwtSettings:RefreshTokenExpiryDays"] ?? "7");
        }

        public async Task<User?> RegisterAsync(string email, string username, string password, string? role = null, string? fullName = null)
        {
            // Kiểm tra user đã tồn tại
            var existingUser = await _users.Find(u => u.Username == username || u.Email == email).FirstOrDefaultAsync();
            if (existingUser != null)
                return null;

            // Validate role
            var userRole = role ?? UserRoles.Teacher;
            if (!UserRoles.IsValidRole(userRole))
                userRole = UserRoles.Student;

            // Hash password
            var passwordHash = BCrypt.Net.BCrypt.HashPassword(password);

            // Lấy permissions theo role
            var permissions = Permissions.GetRolePermissions().ContainsKey(userRole)
                ? Permissions.GetRolePermissions()[userRole]
                : new List<string>();

            var newUser = new User
            {
                Email = email,
                Username = username,
                PasswordHash = passwordHash,
                Role = userRole,
                Permissions = permissions,
                FullName = fullName,
                CreatedAt = DateTime.UtcNow,
                IsActive = true
            };

            await _users.InsertOneAsync(newUser);
            return newUser;
        }

        public async Task<AuthResponse?> LoginAsync(string username, string password)
        {
            var user = await _users.Find(u => u.Username == username).FirstOrDefaultAsync();

            if (user == null || !BCrypt.Net.BCrypt.Verify(password, user.PasswordHash))
                return null;

            // Kiểm tra tài khoản có active không
            if (!user.IsActive)
                return null;

            // 1. Tạo Access Token (JWT)
            var accessToken = GenerateJwtToken(user);

            // 2. Tạo Refresh Token
            var refreshToken = GenerateRefreshToken();
            var refreshTokenExpiry = DateTime.UtcNow.AddDays(_refreshTokenExpiryDays);

            // 3. Lưu Refresh Token vào User trong DB
            var update = Builders<User>.Update
                .Set(u => u.LastLogin, DateTime.UtcNow)
                .Set(u => u.RefreshToken, refreshToken)
                .Set(u => u.RefreshTokenExpiry, refreshTokenExpiry);
            await _users.UpdateOneAsync(u => u.Id == user.Id, update);

            // 4. Trả về cả hai token
            return new AuthResponse
            {
                AccessToken = accessToken,
                RefreshToken = refreshToken,
                UserId = user.Id ?? string.Empty,
                Username = user.Username,
                Email = user.Email,
                Role = user.Role,
                Permissions = user.Permissions,
                FullName = user.FullName,
                Avatar = user.Avatar
            };
        }

        public async Task<AuthResponse?> RefreshTokenAsync(string refreshToken)
        {
            // Tìm user bằng refresh token
            var user = await _users.Find(u => u.RefreshToken == refreshToken).FirstOrDefaultAsync();

            if (user == null)
                return null; // Refresh token không hợp lệ

            // Kiểm tra refresh token có hết hạn không
            if (user.RefreshTokenExpiry.HasValue && user.RefreshTokenExpiry.Value < DateTime.UtcNow)
                return null; // Refresh token đã hết hạn

            // Kiểm tra tài khoản có active không
            if (!user.IsActive)
                return null;

            // Tạo access token và refresh token mới (xoay vòng token để tăng bảo mật)
            var newAccessToken = GenerateJwtToken(user);
            var newRefreshToken = GenerateRefreshToken();
            var newRefreshTokenExpiry = DateTime.UtcNow.AddDays(_refreshTokenExpiryDays);

            // Cập nhật refresh token mới cho user
            var update = Builders<User>.Update
                .Set(u => u.RefreshToken, newRefreshToken)
                .Set(u => u.RefreshTokenExpiry, newRefreshTokenExpiry);
            await _users.UpdateOneAsync(u => u.Id == user.Id, update);

            return new AuthResponse
            {
                AccessToken = newAccessToken,
                RefreshToken = newRefreshToken,
                UserId = user.Id ?? string.Empty,
                Username = user.Username,
                Email = user.Email,
                Role = user.Role,
                Permissions = user.Permissions,
                FullName = user.FullName,
                Avatar = user.Avatar
            };
        }

        public async Task<bool> RevokeRefreshTokenAsync(string userId)
        {
            var update = Builders<User>.Update
                .Set(u => u.RefreshToken, null)
                .Set(u => u.RefreshTokenExpiry, null);

            var result = await _users.UpdateOneAsync(u => u.Id == userId, update);
            return result.ModifiedCount > 0;
        }

        public async Task<User?> GetUserByIdAsync(string id)
        {
            return await _users.Find(u => u.Id == id).FirstOrDefaultAsync();
        }

        public async Task<User?> GetUserByUsernameAsync(string username)
        {
            return await _users.Find(u => u.Username == username).FirstOrDefaultAsync();
        }

        private string GenerateJwtToken(User user)
        {
            var tokenHandler = new JwtSecurityTokenHandler();
            var key = Encoding.ASCII.GetBytes(_jwtSecretKey);

            var claims = new List<Claim>
            {
                new Claim(ClaimTypes.NameIdentifier, user.Id ?? string.Empty),
                new Claim(ClaimTypes.Name, user.Username),
                new Claim(ClaimTypes.Email, user.Email),
                new Claim(ClaimTypes.Role, user.Role)
            };

            // Thêm permissions vào claims
            foreach (var permission in user.Permissions)
            {
                claims.Add(new Claim("permission", permission));
            }

            var tokenDescriptor = new SecurityTokenDescriptor
            {
                Subject = new ClaimsIdentity(claims),
                Expires = DateTime.UtcNow.AddMinutes(_jwtExpiryMinutes),
                Issuer = _jwtIssuer,
                Audience = _jwtAudience,
                SigningCredentials = new SigningCredentials(
                    new SymmetricSecurityKey(key),
                    SecurityAlgorithms.HmacSha256Signature)
            };

            var token = tokenHandler.CreateToken(tokenDescriptor);
            return tokenHandler.WriteToken(token);
        }

        private string GenerateRefreshToken()
        {
            var randomNumber = new byte[32];
            using (var rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(randomNumber);
                return Convert.ToBase64String(randomNumber);
            }
        }
    }
}
