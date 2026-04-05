using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Services;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.API.Middlewares;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using System.Text;
using MongoDB.Driver;

var builder = WebApplication.CreateBuilder(args);
var appMode = NormalizeAppMode(builder.Configuration["AppMode"], builder.Environment);

// Keep local logging on console/debug to avoid EventLog permission failures on Windows.
builder.Logging.ClearProviders();
builder.Logging.AddConfiguration(builder.Configuration.GetSection("Logging"));
builder.Logging.AddConsole();
builder.Logging.AddDebug();
builder.Logging.AddEventSourceLogger();

// ========== CẤU HÌNH MONGODB ==========
builder.Services.Configure<MongoDbSettings>(
    builder.Configuration.GetSection("MongoDbSettings"));
builder.Services.Configure<GoogleSheetsSettings>(
    builder.Configuration.GetSection("GoogleSheets"));
builder.Services.Configure<RedisSettings>(
    builder.Configuration.GetSection("Redis"));

if (appMode == "local")
{
    var dataProtectionKeyPath = Path.Combine(builder.Environment.ContentRootPath, ".data-protection-keys");
    builder.Services.AddDataProtection()
        .SetApplicationName("MOS.ExcelGrading.Local")
        .PersistKeysToFileSystem(new DirectoryInfo(dataProtectionKeyPath));
}

var redisSettings = builder.Configuration.GetSection("Redis").Get<RedisSettings>() ?? new RedisSettings();
var useRedisCache = redisSettings.Enabled && !string.IsNullOrWhiteSpace(redisSettings.ConnectionString);

if (useRedisCache)
{
    builder.Services.AddStackExchangeRedisCache(options =>
    {
        options.Configuration = redisSettings.ConnectionString;
        options.InstanceName = string.IsNullOrWhiteSpace(redisSettings.InstanceName)
            ? "mos-grader:"
            : $"{redisSettings.InstanceName.Trim()}:";
    });
}
else
{
    // Fallback để app vẫn chạy kể cả khi chưa cấu hình Redis
    builder.Services.AddDistributedMemoryCache();
}

// ✅ ĐĂNG KÝ IMongoDatabase
builder.Services.AddSingleton<IMongoDatabase>(sp =>
{
    var settings = builder.Configuration.GetSection("MongoDbSettings").Get<MongoDbSettings>()
        ?? throw new InvalidOperationException("Thiếu cấu hình MongoDbSettings");

    if (string.IsNullOrWhiteSpace(settings.ConnectionString))
        throw new InvalidOperationException(
            "Thiếu MongoDbSettings.ConnectionString. " +
            "Hãy đặt trong appsettings.Development.json hoặc biến môi trường MongoDbSettings__ConnectionString.");

    if (string.IsNullOrWhiteSpace(settings.DatabaseName))
        throw new InvalidOperationException("Thiếu MongoDbSettings.DatabaseName");

    var client = new MongoClient(settings.ConnectionString);
    return client.GetDatabase(settings.DatabaseName);
});

// ========== ĐĂNG KÝ SERVICES ==========
builder.Services.AddScoped<IGradingService, GradingService>();
builder.Services.AddScoped<IAnalyticsService, AnalyticsService>();
builder.Services.AddScoped<IUserService, UserService>();
builder.Services.AddScoped<ISchoolService, SchoolService>();
builder.Services.AddScoped<IClassService, ClassService>();
builder.Services.AddScoped<IStudentService, StudentService>();
builder.Services.AddScoped<IAssignmentService, AssignmentService>();
builder.Services.AddScoped<IScoreService, ScoreService>();
builder.Services.AddScoped<IGradingTestBugNoteService, GradingTestBugNoteService>();
builder.Services.AddScoped<IScheduleService, ScheduleService>();
builder.Services.AddScoped<IAttendanceService, AttendanceService>();
builder.Services.AddScoped<IGoogleSheetAttendanceSyncService, GoogleSheetAttendanceSyncService>();
builder.Services.AddScoped<IComputerRoomService, ComputerRoomService>();
// ========== CẤU HÌNH JWT AUTHENTICATION ==========
var jwtSettings = builder.Configuration.GetSection("JwtSettings");
var secretKey = jwtSettings["SecretKey"] ?? throw new ArgumentNullException("JWT SecretKey không được để trống");

builder.Services.AddAuthentication(options =>
{
    options.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
    options.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
})
.AddJwtBearer(options =>
{
    options.TokenValidationParameters = new TokenValidationParameters
    {
        ValidateIssuer = true,
        ValidateAudience = true,
        ValidateLifetime = true,
        ValidateIssuerSigningKey = true,
        ValidIssuer = jwtSettings["Issuer"],
        ValidAudience = jwtSettings["Audience"],
        IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(secretKey)),
        ClockSkew = TimeSpan.Zero
    };
});

// ========== TĂNG GIỚI HẠN UPLOAD LÊN 500MB ==========
builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 524288000; // 500MB
    options.ValueLengthLimit = int.MaxValue;
    options.MultipartHeadersLengthLimit = int.MaxValue;
});

// ========== CẤU HÌNH CORS ==========
var corsProfile = appMode;

var allowedOrigins = builder.Configuration
    .GetSection($"Cors:Profiles:{corsProfile}")
    .Get<string[]>() ?? Array.Empty<string>();

if (allowedOrigins.Length == 0)
{
    // Backward-compatible fallback
    allowedOrigins = builder.Configuration
        .GetSection("Cors:AllowedOrigins")
        .Get<string[]>() ?? Array.Empty<string>();
}

builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowConfiguredOrigins", policy =>
    {
        if (allowedOrigins.Length == 0)
        {
            policy.WithOrigins("http://localhost:5173", "https://localhost:5173")
                  .AllowAnyMethod()
                  .AllowAnyHeader();
        }
        else
        {
            policy.WithOrigins(allowedOrigins)
                  .AllowAnyMethod()
                  .AllowAnyHeader();
        }
    });
});

// ========== CẤU HÌNH CONTROLLERS ==========
builder.Services.AddControllers();

var app = builder.Build();
app.Logger.LogInformation("AppMode: {AppMode}", appMode);
app.Logger.LogInformation("CorsProfile: {CorsProfile}", corsProfile);
app.Logger.LogInformation("CORS allowed origins: {Origins}", string.Join(", ", allowedOrigins));
app.Logger.LogInformation(
    "Cache provider: {Provider}",
    useRedisCache ? "Redis (StackExchangeRedisCache)" : "InMemory (DistributedMemoryCache)");

// ========== MIDDLEWARE PIPELINE ==========
app.UseCors("AllowConfiguredOrigins");
if (appMode == "local")
{
    app.UseHttpsRedirection();
}
app.UseMiddleware<RequestLoggingMiddleware>();
app.UseAuthentication();
app.UseAuthorization();
app.MapControllers();

app.Run();

static string NormalizeAppMode(string? configuredValue, IHostEnvironment environment)
{
    var mode = configuredValue?.Trim().ToLowerInvariant();

    if (string.IsNullOrWhiteSpace(mode))
    {
        return environment.IsDevelopment() ? "local" : "deploy";
    }

    return mode switch
    {
        "local" => "local",
        "deploy" => "deploy",
        "production" => "deploy",
        "prod" => "deploy",
        _ => environment.IsDevelopment() ? "local" : "deploy"
    };
}
