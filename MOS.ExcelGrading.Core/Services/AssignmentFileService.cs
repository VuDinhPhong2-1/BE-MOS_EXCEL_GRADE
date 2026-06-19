using System.Security.Cryptography;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB.Driver.GridFS;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class AssignmentFileService : IAssignmentFileService
    {
        private const long MaxUploadBytes = 100 * 1024 * 1024; // 100MB mỗi file

        private static readonly HashSet<string> ExcelExtensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".xlsx", ".xlsm", ".xls"
        };

        private static readonly HashSet<string> WordExtensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".docx", ".doc"
        };

        private static readonly HashSet<string> PptExtensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".pptx", ".ppt"
        };

        private static readonly HashSet<string> TextExtensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".txt"
        };

        private readonly IMongoCollection<AssignmentFile> _assignmentFiles;
        private readonly IMongoCollection<Assignment> _assignments;
        private readonly GridFSBucket _bucket;
        private readonly ILogger<AssignmentFileService> _logger;

        public AssignmentFileService(IMongoDatabase database, ILogger<AssignmentFileService> logger)
        {
            _assignmentFiles = database.GetCollection<AssignmentFile>("assignment_files");
            _assignments = database.GetCollection<Assignment>("assignments");
            _bucket = new GridFSBucket(database, new GridFSBucketOptions
            {
                BucketName = "assignmentFiles"
            });
            _logger = logger;
        }

        public async Task<AssignmentFileResponse> UploadAssignmentFileAsync(
            string assignmentId,
            IFormFile file,
            string subject,
            string kind,
            string userId)
        {
            EnsureValidObjectId(assignmentId, "assignmentId");
            EnsureValidObjectId(userId, "userId");

            if (file == null)
            {
                throw new ArgumentException("Cần cung cấp file upload.");
            }

            if (file.Length <= 0)
            {
                throw new ArgumentException("File rỗng hoặc không hợp lệ.");
            }

            if (file.Length > MaxUploadBytes)
            {
                throw new ArgumentException($"Dung lượng file vượt quá giới hạn {MaxUploadBytes / (1024 * 1024)}MB.");
            }

            var normalizedSubject = AssignmentFileSubjects.Normalize(subject);
            if (!AssignmentFileSubjects.IsValid(normalizedSubject))
            {
                throw new ArgumentException("subject không hợp lệ. Chỉ chấp nhận: excel, word, ppt.");
            }

            var normalizedKind = AssignmentFileKinds.Normalize(kind);
            if (!AssignmentFileKinds.IsValid(normalizedKind))
            {
                throw new ArgumentException("kind không hợp lệ. Chỉ chấp nhận: template hoặc answer.");
            }

            var assignment = await _assignments
                .Find(a => a.Id == assignmentId && a.IsActive)
                .FirstOrDefaultAsync();
            if (assignment == null)
            {
                throw new InvalidOperationException("Không tìm thấy bài tập hoặc bài tập đã ngừng hoạt động.");
            }

            var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
            ValidateExtension(normalizedSubject, normalizedKind, extension);

            await using var inMemory = new MemoryStream();
            await file.CopyToAsync(inMemory);
            if (inMemory.Length <= 0)
            {
                throw new ArgumentException("Nội dung file không hợp lệ.");
            }

            var fileBytes = inMemory.ToArray();
            var sha256 = Convert.ToHexString(SHA256.HashData(fileBytes)).ToLowerInvariant();

            var duplicatedActive = await _assignmentFiles
                .Find(f =>
                    f.AssignmentId == assignmentId &&
                    f.Subject == normalizedSubject &&
                    f.Kind == normalizedKind &&
                    f.Sha256 == sha256 &&
                    f.IsActive)
                .FirstOrDefaultAsync();
            if (duplicatedActive != null)
            {
                return ToResponse(duplicatedActive);
            }

            var latestVersion = await _assignmentFiles
                .Find(f =>
                    f.AssignmentId == assignmentId &&
                    f.Subject == normalizedSubject &&
                    f.Kind == normalizedKind)
                .SortByDescending(f => f.Version)
                .FirstOrDefaultAsync();

            var nextVersion = (latestVersion?.Version ?? 0) + 1;
            var now = DateTime.UtcNow;
            inMemory.Position = 0;

            var uploadMetadata = new BsonDocument
            {
                { "assignmentId", assignmentId },
                { "classId", assignment.ClassId },
                { "subject", normalizedSubject },
                { "kind", normalizedKind },
                { "version", nextVersion },
                { "sha256", sha256 },
                { "originalName", file.FileName },
                { "uploadedBy", userId },
                { "uploadedAt", now }
            };

            var gridFsFileId = await _bucket.UploadFromStreamAsync(
                file.FileName,
                inMemory,
                new GridFSUploadOptions
                {
                    Metadata = uploadMetadata
                });

            // Mỗi assignment + subject + kind chỉ có 1 bản active hiện hành.
            var deactivateUpdate = Builders<AssignmentFile>.Update
                .Set(x => x.IsActive, false)
                .Set(x => x.UpdatedAt, now)
                .Set(x => x.UpdatedBy, userId);
            await _assignmentFiles.UpdateManyAsync(
                x =>
                    x.AssignmentId == assignmentId &&
                    x.Subject == normalizedSubject &&
                    x.Kind == normalizedKind &&
                    x.IsActive,
                deactivateUpdate);

            var uploadedFile = new AssignmentFile
            {
                AssignmentId = assignmentId,
                ClassId = assignment.ClassId,
                GridFsFileId = gridFsFileId.ToString(),
                OriginalName = file.FileName,
                Extension = extension,
                ContentType = ResolveContentType(file.ContentType, extension),
                SizeBytes = file.Length,
                Sha256 = sha256,
                Subject = normalizedSubject,
                Kind = normalizedKind,
                Version = nextVersion,
                IsActive = true,
                UploadedAt = now,
                UploadedBy = userId
            };

            await _assignmentFiles.InsertOneAsync(uploadedFile);

            var assignmentUpdates = new List<UpdateDefinition<Assignment>>
            {
                Builders<Assignment>.Update.Set(a => a.UpdatedAt, now),
                Builders<Assignment>.Update.Set(a => a.UpdatedBy, userId)
            };

            if (normalizedKind == AssignmentFileKinds.Template)
            {
                assignmentUpdates.Add(
                    Builders<Assignment>.Update.Set(a => a.CurrentTemplateFileId, uploadedFile.Id));
            }
            else if (normalizedKind == AssignmentFileKinds.Answer)
            {
                assignmentUpdates.Add(
                    Builders<Assignment>.Update.Set(a => a.CurrentAnswerFileId, uploadedFile.Id));
            }
            else if (normalizedKind == AssignmentFileKinds.Instructions)
            {
                assignmentUpdates.Add(
                    Builders<Assignment>.Update.Set(a => a.CurrentInstructionsFileId, uploadedFile.Id));
            }
            else if (normalizedKind == AssignmentFileKinds.Help)
            {
                assignmentUpdates.Add(
                    Builders<Assignment>.Update.Set(a => a.CurrentHelpFileId, uploadedFile.Id));
            }

            await _assignments.UpdateOneAsync(
                a => a.Id == assignmentId,
                Builders<Assignment>.Update.Combine(assignmentUpdates));

            _logger.LogInformation(
                "✅ Uploaded assignment file: {FileName} | AssignmentId: {AssignmentId} | Subject: {Subject} | Kind: {Kind} | Version: {Version}",
                file.FileName,
                assignmentId,
                normalizedSubject,
                normalizedKind,
                nextVersion);

            return ToResponse(uploadedFile);
        }

        public async Task<List<AssignmentFileResponse>> GetAssignmentFilesAsync(
            string assignmentId,
            bool includeInactive = false)
        {
            EnsureValidObjectId(assignmentId, "assignmentId");

            var filter = Builders<AssignmentFile>.Filter.Eq(x => x.AssignmentId, assignmentId);
            if (!includeInactive)
            {
                filter &= Builders<AssignmentFile>.Filter.Eq(x => x.IsActive, true);
            }

            var files = await _assignmentFiles
                .Find(filter)
                .SortByDescending(x => x.UploadedAt)
                .ThenByDescending(x => x.Version)
                .ToListAsync();

            return files.Select(ToResponse).ToList();
        }

        public async Task<AssignmentFileResponse?> GetAssignmentFileByIdAsync(string fileId)
        {
            EnsureValidObjectId(fileId, "fileId");
            var file = await _assignmentFiles.Find(x => x.Id == fileId).FirstOrDefaultAsync();
            return file == null ? null : ToResponse(file);
        }

        public async Task<AssignmentFileDownloadResult> OpenAssignmentFileDownloadAsync(string fileId)
        {
            EnsureValidObjectId(fileId, "fileId");

            var file = await _assignmentFiles.Find(x => x.Id == fileId).FirstOrDefaultAsync();
            if (file == null)
            {
                throw new InvalidOperationException("Không tìm thấy file bài tập.");
            }

            if (!ObjectId.TryParse(file.GridFsFileId, out var gridFsObjectId))
            {
                throw new InvalidOperationException("File metadata không hợp lệ (gridFsFileId).");
            }

            var stream = await _bucket.OpenDownloadStreamAsync(gridFsObjectId);
            return new AssignmentFileDownloadResult
            {
                Stream = stream,
                FileName = file.OriginalName,
                ContentType = ResolveContentType(file.ContentType, file.Extension)
            };
        }

        public async Task<bool> SoftDeleteAssignmentFileAsync(string fileId, string userId)
        {
            EnsureValidObjectId(fileId, "fileId");
            EnsureValidObjectId(userId, "userId");

            var file = await _assignmentFiles.Find(x => x.Id == fileId).FirstOrDefaultAsync();
            if (file == null)
            {
                return false;
            }

            if (!file.IsActive)
            {
                return true;
            }

            var now = DateTime.UtcNow;
            var fileUpdate = Builders<AssignmentFile>.Update
                .Set(x => x.IsActive, false)
                .Set(x => x.UpdatedAt, now)
                .Set(x => x.UpdatedBy, userId);

            await _assignmentFiles.UpdateOneAsync(x => x.Id == fileId, fileUpdate);

            var assignmentUpdates = new List<UpdateDefinition<Assignment>>
            {
                Builders<Assignment>.Update.Set(a => a.UpdatedAt, now),
                Builders<Assignment>.Update.Set(a => a.UpdatedBy, userId)
            };

            if (file.Kind == AssignmentFileKinds.Template)
            {
                assignmentUpdates.Add(Builders<Assignment>.Update.Set(a => a.CurrentTemplateFileId, null));
            }
            else if (file.Kind == AssignmentFileKinds.Answer)
            {
                assignmentUpdates.Add(Builders<Assignment>.Update.Set(a => a.CurrentAnswerFileId, null));
            }
            else if (file.Kind == AssignmentFileKinds.Instructions)
            {
                assignmentUpdates.Add(Builders<Assignment>.Update.Set(a => a.CurrentInstructionsFileId, null));
            }
            else if (file.Kind == AssignmentFileKinds.Help)
            {
                assignmentUpdates.Add(Builders<Assignment>.Update.Set(a => a.CurrentHelpFileId, null));
            }

            await _assignments.UpdateOneAsync(
                a => a.Id == file.AssignmentId,
                Builders<Assignment>.Update.Combine(assignmentUpdates));

            return true;
        }

        private static AssignmentFileResponse ToResponse(AssignmentFile file) =>
            new()
            {
                Id = file.Id,
                AssignmentId = file.AssignmentId,
                ClassId = file.ClassId,
                OriginalName = file.OriginalName,
                Extension = file.Extension,
                ContentType = file.ContentType,
                SizeBytes = file.SizeBytes,
                Sha256 = file.Sha256,
                Subject = file.Subject,
                Kind = file.Kind,
                Version = file.Version,
                IsActive = file.IsActive,
                UploadedAt = file.UploadedAt,
                UploadedBy = file.UploadedBy,
                UpdatedAt = file.UpdatedAt,
                UpdatedBy = file.UpdatedBy
            };

        private static void EnsureValidObjectId(string value, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(value) || !ObjectId.TryParse(value, out _))
            {
                throw new ArgumentException($"{fieldName} không hợp lệ.");
            }
        }

        private static void ValidateExtension(string subject, string kind, string extension)
        {
            if (kind == AssignmentFileKinds.Instructions || kind == AssignmentFileKinds.Help)
            {
                if (!TextExtensions.Contains(extension))
                {
                    throw new ArgumentException(
                        $"Dinh dang file '{extension}' khong hop le cho loai '{kind}'. Chi chap nhan file .txt.");
                }

                return;
            }

            var allowed = subject switch
            {
                AssignmentFileSubjects.Excel => ExcelExtensions,
                AssignmentFileSubjects.Word => WordExtensions,
                AssignmentFileSubjects.Ppt => PptExtensions,
                _ => throw new ArgumentException("subject không hợp lệ.")
            };

            if (!allowed.Contains(extension))
            {
                throw new ArgumentException(
                    $"Định dạng file '{extension}' không hợp lệ cho môn '{subject}'.");
            }
        }

        private static string ResolveContentType(string? requestedContentType, string extension)
        {
            if (!string.IsNullOrWhiteSpace(requestedContentType) &&
                !requestedContentType.Equals("application/octet-stream", StringComparison.OrdinalIgnoreCase))
            {
                return requestedContentType;
            }

            return extension.ToLowerInvariant() switch
            {
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ".xlsm" => "application/vnd.ms-excel.sheet.macroEnabled.12",
                ".xls" => "application/vnd.ms-excel",
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".doc" => "application/msword",
                ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                ".ppt" => "application/vnd.ms-powerpoint",
                _ => "application/octet-stream"
            };
        }
    }
}
