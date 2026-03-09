// Core/Services/UserService.cs
using Google.Apis.Auth;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Tokens;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;

namespace MOS.ExcelGrading.Core.Services
{
    public class UserService : IUserService
    {
        private readonly IMongoCollection<User> _users;
        private static int _indexInitialized;
        private readonly string _jwtSecretKey;
        private readonly string _jwtIssuer;
        private readonly string _jwtAudience;
        private readonly int _jwtExpiryMinutes;
        private readonly int _refreshTokenExpiryDays;
        private readonly string? _googleClientId;

        public UserService(
            IMongoDatabase database,
            IConfiguration configuration)
        {
            _users = database.GetCollection<User>("Users");

            if (Interlocked.Exchange(ref _indexInitialized, 1) == 0)
            {
                var usernameIndex = new CreateIndexModel<User>(
                    Builders<User>.IndexKeys.Ascending(u => u.Username),
                    new CreateIndexOptions { Unique = true });

                var emailIndex = new CreateIndexModel<User>(
                    Builders<User>.IndexKeys.Ascending(u => u.Email),
                    new CreateIndexOptions { Unique = true });

                var googleIdIndex = new CreateIndexModel<User>(
                    Builders<User>.IndexKeys.Ascending(u => u.GoogleId),
                    new CreateIndexOptions { Unique = true, Sparse = true });

                _users.Indexes.CreateOne(usernameIndex);
                _users.Indexes.CreateOne(emailIndex);
                _users.Indexes.CreateOne(googleIdIndex);
            }

            _jwtSecretKey = configuration["JwtSettings:SecretKey"] ?? throw new ArgumentNullException("JWT SecretKey");
            _jwtIssuer = configuration["JwtSettings:Issuer"] ?? "MOS.ExcelGrading";
            _jwtAudience = configuration["JwtSettings:Audience"] ?? "MOS.ExcelGrading.Users";
            _jwtExpiryMinutes = int.Parse(configuration["JwtSettings:ExpiryMinutes"] ?? "60");
            _refreshTokenExpiryDays = int.Parse(configuration["JwtSettings:RefreshTokenExpiryDays"] ?? "7");
            _googleClientId = configuration["GoogleAuth:ClientId"];
        }

        public async Task<User?> RegisterAsync(string email, string username, string password, string? role = null, string? fullName = null)
        {
            var existingUser = await _users.Find(u => u.Username == username || u.Email == email).FirstOrDefaultAsync();
            if (existingUser != null)
                return null;

            var userRole = role ?? UserRoles.Teacher;
            if (!UserRoles.IsValidRole(userRole))
                userRole = UserRoles.Teacher;

            var passwordHash = BCrypt.Net.BCrypt.HashPassword(password);

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
                IsActive = true,
                AuthProvider = "Local"
            };

            await _users.InsertOneAsync(newUser);
            return newUser;
        }

        public async Task<AuthResponse?> LoginAsync(string username, string password)
        {
            var user = await _users.Find(u => u.Username == username).FirstOrDefaultAsync();

            if (user == null || !BCrypt.Net.BCrypt.Verify(password, user.PasswordHash))
                return null;

            if (!user.IsActive)
                return null;

            return await CreateAuthResponseAsync(user);
        }

        public async Task<AuthResponse?> LoginWithGoogleAsync(string idToken)
        {
            if (string.IsNullOrWhiteSpace(idToken))
                return null;

            if (string.IsNullOrWhiteSpace(_googleClientId))
                throw new InvalidOperationException("Thiếu cấu hình GoogleAuth:ClientId");

            var payload = await GoogleJsonWebSignature.ValidateAsync(idToken, new GoogleJsonWebSignature.ValidationSettings
            {
                Audience = new[] { _googleClientId }
            });

            if (string.IsNullOrWhiteSpace(payload.Email) || payload.EmailVerified != true)
                return null;

            var filter = Builders<User>.Filter.Or(
                Builders<User>.Filter.Eq(u => u.GoogleId, payload.Subject),
                Builders<User>.Filter.Eq(u => u.Email, payload.Email));

            var user = await _users.Find(filter).FirstOrDefaultAsync();

            if (user == null)
            {
                var role = UserRoles.Teacher;
                var newUser = new User
                {
                    Email = payload.Email,
                    Username = await GenerateUniqueUsernameAsync(payload.Email),
                    PasswordHash = BCrypt.Net.BCrypt.HashPassword(Guid.NewGuid().ToString("N")),
                    Role = role,
                    Permissions = Permissions.GetRolePermissions()[role],
                    FullName = payload.Name,
                    Avatar = payload.Picture,
                    GoogleId = payload.Subject,
                    AuthProvider = "Google",
                    IsActive = true,
                    IsEmailVerified = true,
                    CreatedAt = DateTime.UtcNow
                };

                await _users.InsertOneAsync(newUser);
                return await CreateAuthResponseAsync(newUser);
            }

            if (!user.IsActive)
                return null;

            var profileUpdate = Builders<User>.Update
                .Set(u => u.GoogleId, payload.Subject)
                .Set(u => u.AuthProvider, "Google")
                .Set(u => u.IsEmailVerified, true)
                .Set(u => u.Avatar, string.IsNullOrWhiteSpace(payload.Picture) ? user.Avatar : payload.Picture)
                .Set(u => u.FullName, string.IsNullOrWhiteSpace(payload.Name) ? user.FullName : payload.Name);

            user.GoogleId = payload.Subject;
            user.AuthProvider = "Google";
            user.IsEmailVerified = true;
            user.Avatar = string.IsNullOrWhiteSpace(payload.Picture) ? user.Avatar : payload.Picture;
            user.FullName = string.IsNullOrWhiteSpace(payload.Name) ? user.FullName : payload.Name;

            return await CreateAuthResponseAsync(user, profileUpdate);
        }

        public async Task<AuthResponse?> RefreshTokenAsync(string refreshToken)
        {
            var user = await _users.Find(u => u.RefreshToken == refreshToken).FirstOrDefaultAsync();

            if (user == null)
                return null;

            if (user.RefreshTokenExpiry.HasValue && user.RefreshTokenExpiry.Value < DateTime.UtcNow)
                return null;

            if (!user.IsActive)
                return null;

            var newAccessToken = GenerateJwtToken(user);
            var newRefreshToken = GenerateRefreshToken();
            var newRefreshTokenExpiry = DateTime.UtcNow.AddDays(_refreshTokenExpiryDays);

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

        public async Task<User?> UpdateProfileAsync(string userId, UpdateProfileRequest request)
        {
            var fullName = string.IsNullOrWhiteSpace(request.FullName) ? null : request.FullName.Trim();
            var phoneNumber = string.IsNullOrWhiteSpace(request.PhoneNumber) ? null : request.PhoneNumber.Trim();
            var avatar = string.IsNullOrWhiteSpace(request.Avatar) ? null : request.Avatar.Trim();

            var update = Builders<User>.Update
                .Set(u => u.FullName, fullName)
                .Set(u => u.PhoneNumber, phoneNumber)
                .Set(u => u.Avatar, avatar);

            return await _users.FindOneAndUpdateAsync(
                filter: u => u.Id == userId && u.IsActive,
                update: update,
                options: new FindOneAndUpdateOptions<User>
                {
                    ReturnDocument = ReturnDocument.After
                });
        }

        private async Task<string> GenerateUniqueUsernameAsync(string email)
        {
            var baseUsername = email.Split('@')[0];
            var sanitized = new string(baseUsername
                .Where(c => char.IsLetterOrDigit(c) || c == '_' || c == '.')
                .ToArray())
                .Trim();

            if (string.IsNullOrWhiteSpace(sanitized))
                sanitized = "user";

            if (sanitized.Length < 3)
                sanitized = sanitized.PadRight(3, '0');

            var candidate = sanitized;
            var suffix = 1;

            while (await _users.Find(u => u.Username == candidate).AnyAsync())
            {
                candidate = $"{sanitized}{suffix}";
                suffix++;
            }

            return candidate;
        }

        private async Task<AuthResponse> CreateAuthResponseAsync(User user, UpdateDefinition<User>? additionalUpdate = null)
        {
            var accessToken = GenerateJwtToken(user);
            var refreshToken = GenerateRefreshToken();
            var refreshTokenExpiry = DateTime.UtcNow.AddDays(_refreshTokenExpiryDays);

            var update = Builders<User>.Update
                .Set(u => u.LastLogin, DateTime.UtcNow)
                .Set(u => u.RefreshToken, refreshToken)
                .Set(u => u.RefreshTokenExpiry, refreshTokenExpiry);

            if (additionalUpdate != null)
            {
                update = Builders<User>.Update.Combine(update, additionalUpdate);
            }

            await _users.UpdateOneAsync(u => u.Id == user.Id, update);

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
            using var rng = RandomNumberGenerator.Create();
            rng.GetBytes(randomNumber);
            return Convert.ToBase64String(randomNumber);
        }
    }
}

