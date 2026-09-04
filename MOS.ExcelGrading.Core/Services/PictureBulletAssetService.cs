using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB.Driver.GridFS;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Utilities;

namespace MOS.ExcelGrading.Core.Services
{
    /// <summary>
    /// Lưu ảnh bullet chuẩn vào MongoDB GridFS (dùng chung cluster với
    /// grading_rule_sets, không cần thêm hạ tầng storage riêng).
    /// </summary>
    public class PictureBulletAssetService : IImageAssetService
    {
        // Khớp đúng giới hạn phía FE (PictureBulletEditor.tsx: MAX_FILE_SIZE)
        private const long MaxFileSizeBytes = 10 * 1024 * 1024;

        // Khớp đúng ACCEPTED_IMAGE_TYPES phía FE
        private static readonly HashSet<string> SupportedContentTypes = new(StringComparer.OrdinalIgnoreCase)
        {
            "image/png",
            "image/jpeg",
            "image/gif",
            "image/bmp",
            "image/webp"
        };

        private readonly GridFSBucket _bucket;

        public PictureBulletAssetService(IMongoDatabase database)
        {
            _bucket = new GridFSBucket(database, new GridFSBucketOptions
            {
                BucketName = "pictureBulletAssets"
            });
        }

        public async Task<PictureBulletAssetUploadResult> UploadAsync(Stream content, string fileName, string contentType, ImageAssetKind kind)
        {
            if (content == null)
            {
                throw new InvalidOperationException("File ảnh không được rỗng.");
            }

            if (string.IsNullOrWhiteSpace(contentType) || !SupportedContentTypes.Contains(contentType))
            {
                throw new InvalidOperationException("Định dạng ảnh không được hỗ trợ. Chỉ chấp nhận PNG, JPG, GIF, BMP hoặc WebP.");
            }

            using var memoryStream = new MemoryStream();
            await content.CopyToAsync(memoryStream);
            var bytes = memoryStream.ToArray();

            if (bytes.LongLength == 0)
            {
                throw new InvalidOperationException("File ảnh không được rỗng.");
            }

            if (bytes.LongLength > MaxFileSizeBytes)
            {
                throw new InvalidOperationException("Kích thước ảnh không được vượt quá 10MB.");
            }

            var imageHash = ImageHashUtility.ComputeSha256(bytes);

            var uploadOptions = new GridFSUploadOptions
            {
                Metadata = new BsonDocument
                {
                    { "contentType", contentType },
                    { "imageHash", imageHash },
                    { "kind", kind.ToString() },
                    { "uploadedAtUtc", DateTime.UtcNow }
                }
            };

            memoryStream.Position = 0;
            var objectId = await _bucket.UploadFromStreamAsync(
                string.IsNullOrWhiteSpace(fileName) ? "picture-bullet" : fileName,
                memoryStream,
                uploadOptions);

            return new PictureBulletAssetUploadResult
            {
                AssetId = objectId.ToString(),
                ImageHash = imageHash,
                ContentType = contentType,
                SizeBytes = bytes.LongLength
            };
        }

        public async Task<PictureBulletAssetContent?> GetAsync(string assetId)
        {
            if (string.IsNullOrWhiteSpace(assetId) || !ObjectId.TryParse(assetId, out var objectId))
            {
                return null;
            }

            try
            {
                using var downloadStream = await _bucket.OpenDownloadStreamAsync(objectId);

                using var buffer = new MemoryStream();
                await downloadStream.CopyToAsync(buffer);

                var contentType = "application/octet-stream";
                if (downloadStream.FileInfo?.Metadata != null &&
                    downloadStream.FileInfo.Metadata.TryGetValue("contentType", out var contentTypeValue))
                {
                    contentType = contentTypeValue.AsString;
                }

                return new PictureBulletAssetContent
                {
                    Content = buffer.ToArray(),
                    ContentType = contentType,
                    FileName = downloadStream.FileInfo?.Filename ?? "picture-bullet"
                };
            }
            catch (GridFSFileNotFoundException)
            {
                return null;
            }
        }
    }
}