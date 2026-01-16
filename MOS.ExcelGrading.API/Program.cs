using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Services;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.API.Middlewares;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using System.Text;
using MongoDB.Driver;  // ✅ THÊM DÒNG NÀY

var builder = WebApplication.CreateBuilder(args);

// ========== CẤU HÌNH MONGODB ==========
builder.Services.Configure<MongoDbSettings>(
    builder.Configuration.GetSection("MongoDbSettings"));

// ✅ ĐĂNG KÝ IMongoDatabase
builder.Services.AddSingleton<IMongoDatabase>(sp =>
{
    var settings = builder.Configuration.GetSection("MongoDbSettings").Get<MongoDbSettings>();
    var client = new MongoClient(settings.ConnectionString);
    return client.GetDatabase(settings.DatabaseName);
});

// ========== ĐĂNG KÝ SERVICES ==========
builder.Services.AddScoped<IGradingService, GradingService>();
builder.Services.AddScoped<IUserService, UserService>();
builder.Services.AddScoped<ISchoolService, SchoolService>();
builder.Services.AddScoped<IClassService, ClassService>();
builder.Services.AddScoped<IStudentService, StudentService>();
builder.Services.AddScoped<IAssignmentService, AssignmentService>();
builder.Services.AddScoped<IScoreService, ScoreService>();
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
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyMethod()
              .AllowAnyHeader();
    });
});

// ========== CẤU HÌNH CONTROLLERS ==========
builder.Services.AddControllers();

var app = builder.Build();

// ========== MIDDLEWARE PIPELINE ==========
app.UseCors("AllowAll");
app.UseHttpsRedirection();
app.UseMiddleware<RequestLoggingMiddleware>();
app.UseAuthentication();
app.UseAuthorization();
app.MapControllers();

app.Run();
