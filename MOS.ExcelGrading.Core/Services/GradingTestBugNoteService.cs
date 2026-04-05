using Microsoft.Extensions.Logging;
using MongoDB.Bson;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Text.RegularExpressions;

namespace MOS.ExcelGrading.Core.Services
{
    public class GradingTestBugNoteService : IGradingTestBugNoteService
    {
        private static readonly Regex ProjectCodeRegex = new(
            @"^project(?<number>\d{1,2})$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly HashSet<string> AllowedSeverities = new(StringComparer.OrdinalIgnoreCase)
        {
            "low",
            "medium",
            "high",
            "critical"
        };

        private readonly IMongoCollection<GradingTestBugNote> _bugNotes;
        private readonly ILogger<GradingTestBugNoteService> _logger;

        public GradingTestBugNoteService(
            IMongoDatabase database,
            ILogger<GradingTestBugNoteService> logger)
        {
            _bugNotes = database.GetCollection<GradingTestBugNote>("gradingTestBugNotes");
            _logger = logger;

            EnsureIndexes();
        }

        public async Task<List<GradingTestBugNoteResponse>> GetByUserAsync(string userId, string? projectCode = null)
        {
            EnsureValidObjectId(userId, "Người dùng");
            var normalizedUserId = userId.Trim();

            var filter = Builders<GradingTestBugNote>.Filter.Eq(note => note.UserId, normalizedUserId);
            if (!string.IsNullOrWhiteSpace(projectCode))
            {
                var normalizedProjectCode = NormalizeProjectCode(projectCode);
                filter = Builders<GradingTestBugNote>.Filter.And(
                    filter,
                    Builders<GradingTestBugNote>.Filter.Eq(note => note.ProjectCode, normalizedProjectCode));
            }

            var notes = await _bugNotes
                .Find(filter)
                .SortByDescending(note => note.CreatedAt)
                .Limit(1000)
                .ToListAsync();

            return notes.Select(ToResponse).ToList();
        }

        public async Task<GradingTestBugNoteResponse> CreateAsync(CreateGradingTestBugNoteRequest request, string userId)
        {
            if (request == null)
            {
                throw new InvalidOperationException("Thiếu dữ liệu bug note.");
            }

            EnsureValidObjectId(userId, "Người dùng");
            var normalizedUserId = userId.Trim();
            var normalizedProjectCode = NormalizeProjectCode(request.ProjectCode);

            var note = new GradingTestBugNote
            {
                UserId = normalizedUserId,
                ProjectCode = normalizedProjectCode,
                ProjectDisplayName = NormalizeProjectDisplayName(request.ProjectDisplayName, normalizedProjectCode),
                Title = NormalizeRequiredText(request.Title, 200, "Tiêu đề bug"),
                Description = NormalizeRequiredText(request.Description, 5000, "Mô tả bug"),
                Severity = NormalizeSeverity(request.Severity),
                ScoreSummary = NormalizeScoreSummary(request.ScoreSummary),
                GradingError = NormalizeOptionalText(request.GradingError, 2000),
                CreatedAt = DateTime.UtcNow,
                UpdatedAt = DateTime.UtcNow
            };

            await _bugNotes.InsertOneAsync(note);
            _logger.LogInformation(
                "✅ Created grading test bug note {BugNoteId} by user {UserId}",
                note.Id,
                normalizedUserId);

            return ToResponse(note);
        }

        public async Task<bool> DeleteAsync(string id, string userId)
        {
            EnsureValidObjectId(id, "Mã bug note");
            EnsureValidObjectId(userId, "Người dùng");

            var normalizedId = id.Trim();
            var normalizedUserId = userId.Trim();

            var filter = Builders<GradingTestBugNote>.Filter.And(
                Builders<GradingTestBugNote>.Filter.Eq(note => note.Id, normalizedId),
                Builders<GradingTestBugNote>.Filter.Eq(note => note.UserId, normalizedUserId));

            var result = await _bugNotes.DeleteOneAsync(filter);
            if (result.DeletedCount > 0)
            {
                _logger.LogInformation(
                    "✅ Deleted grading test bug note {BugNoteId} by user {UserId}",
                    normalizedId,
                    normalizedUserId);
                return true;
            }

            return false;
        }

        private void EnsureIndexes()
        {
            try
            {
                var indexes = new[]
                {
                    new CreateIndexModel<GradingTestBugNote>(
                        Builders<GradingTestBugNote>.IndexKeys
                            .Ascending(note => note.UserId)
                            .Descending(note => note.CreatedAt)),
                    new CreateIndexModel<GradingTestBugNote>(
                        Builders<GradingTestBugNote>.IndexKeys
                            .Ascending(note => note.UserId)
                            .Ascending(note => note.ProjectCode))
                };

                _bugNotes.Indexes.CreateMany(indexes);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "⚠️ Unable to ensure indexes for grading test bug notes collection");
            }
        }

        private static void EnsureValidObjectId(string? value, string label)
        {
            if (string.IsNullOrWhiteSpace(value) || !ObjectId.TryParse(value.Trim(), out _))
            {
                throw new InvalidOperationException($"{label} không hợp lệ.");
            }
        }

        private static string NormalizeProjectCode(string? projectCode)
        {
            if (string.IsNullOrWhiteSpace(projectCode))
            {
                throw new InvalidOperationException("Project code không được để trống.");
            }

            var normalized = projectCode
                .Trim()
                .Replace("\\", "/", StringComparison.Ordinal)
                .Trim('/')
                .ToLowerInvariant();

            if (normalized.StartsWith("excel/", StringComparison.Ordinal))
            {
                normalized = normalized["excel/".Length..];
            }

            var match = ProjectCodeRegex.Match(normalized);
            if (!match.Success)
            {
                throw new InvalidOperationException("Project code không hợp lệ.");
            }

            if (!int.TryParse(match.Groups["number"].Value, out var projectNumber))
            {
                throw new InvalidOperationException("Project code không hợp lệ.");
            }

            if (projectNumber < 1 || projectNumber > 99)
            {
                throw new InvalidOperationException("Project code không hợp lệ.");
            }

            return $"project{projectNumber:00}";
        }

        private static string NormalizeProjectDisplayName(string? value, string normalizedProjectCode)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                var trimmed = value.Trim();
                return trimmed.Length <= 120 ? trimmed : trimmed[..120];
            }

            var suffix = normalizedProjectCode["project".Length..];
            return int.TryParse(suffix, out var number)
                ? $"Project {number:00} - Excel"
                : normalizedProjectCode;
        }

        private static string NormalizeRequiredText(string? value, int maxLength, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new InvalidOperationException($"{fieldName} không được để trống.");
            }

            var trimmed = value.Trim();
            return trimmed.Length <= maxLength ? trimmed : trimmed[..maxLength];
        }

        private static string? NormalizeOptionalText(string? value, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            var trimmed = value.Trim();
            return trimmed.Length <= maxLength ? trimmed : trimmed[..maxLength];
        }

        private static string NormalizeSeverity(string? severity)
        {
            var normalized = string.IsNullOrWhiteSpace(severity)
                ? "medium"
                : severity.Trim().ToLowerInvariant();

            if (!AllowedSeverities.Contains(normalized))
            {
                throw new InvalidOperationException("Mức độ bug không hợp lệ.");
            }

            return normalized;
        }

        private static GradingTestBugNoteScoreSummary? NormalizeScoreSummary(GradingTestBugNoteScoreSummaryDto? summary)
        {
            if (summary == null)
            {
                return null;
            }

            return new GradingTestBugNoteScoreSummary
            {
                TotalScore = NormalizeFiniteNumber(summary.TotalScore, 0, 100000, "Tổng điểm"),
                MaxScore = NormalizeFiniteNumber(summary.MaxScore, 0, 100000, "Điểm tối đa"),
                Percentage = NormalizeFiniteNumber(summary.Percentage, 0, 1000, "Tỉ lệ điểm"),
                Status = NormalizeOptionalText(summary.Status, 120) ?? string.Empty
            };
        }

        private static double NormalizeFiniteNumber(double value, double min, double max, string fieldName)
        {
            if (!double.IsFinite(value))
            {
                throw new InvalidOperationException($"{fieldName} không hợp lệ.");
            }

            if (value < min || value > max)
            {
                throw new InvalidOperationException($"{fieldName} không hợp lệ.");
            }

            return Math.Round(value, 2, MidpointRounding.AwayFromZero);
        }

        private static GradingTestBugNoteResponse ToResponse(GradingTestBugNote note)
        {
            return new GradingTestBugNoteResponse
            {
                Id = note.Id,
                ProjectCode = note.ProjectCode,
                ProjectDisplayName = note.ProjectDisplayName,
                Title = note.Title,
                Description = note.Description,
                Severity = note.Severity,
                ScoreSummary = note.ScoreSummary == null
                    ? null
                    : new GradingTestBugNoteScoreSummaryDto
                    {
                        TotalScore = note.ScoreSummary.TotalScore,
                        MaxScore = note.ScoreSummary.MaxScore,
                        Percentage = note.ScoreSummary.Percentage,
                        Status = note.ScoreSummary.Status
                    },
                GradingError = note.GradingError,
                CreatedAt = note.CreatedAt
            };
        }
    }
}
