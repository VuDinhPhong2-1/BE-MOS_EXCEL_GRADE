namespace MOS.ExcelGrading.Core.DTOs
{
    public class AssignmentFileResponse
    {
        public string Id { get; set; } = string.Empty;
        public string AssignmentId { get; set; } = string.Empty;
        public string ClassId { get; set; } = string.Empty;
        public string OriginalName { get; set; } = string.Empty;
        public string Extension { get; set; } = string.Empty;
        public string ContentType { get; set; } = "application/octet-stream";
        public long SizeBytes { get; set; }
        public string Sha256 { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string Kind { get; set; } = string.Empty;
        public int Version { get; set; } = 1;
        public bool IsActive { get; set; }
        public DateTime UploadedAt { get; set; }
        public string? UploadedBy { get; set; }
        public DateTime? UpdatedAt { get; set; }
        public string? UpdatedBy { get; set; }
    }

    public class AssignmentFileDownloadResult
    {
        public Stream Stream { get; set; } = Stream.Null;
        public string ContentType { get; set; } = "application/octet-stream";
        public string FileName { get; set; } = string.Empty;
    }
}
