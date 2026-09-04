namespace MOS.ExcelGrading.Core.Models
{
    /// <summary>
    /// Kết quả trả về sau khi upload ảnh bullet chuẩn thành công.
    /// FE sẽ dùng AssetId + ImageHash để điền vào
    /// SpecialCondition.Config (PictureBulletConfig).
    /// </summary>
    public class PictureBulletAssetUploadResult
    {
        public string AssetId { get; set; } = string.Empty;
        public string ImageHash { get; set; } = string.Empty;
        public string PerceptualHash { get; set; } = string.Empty;
        public string ContentType { get; set; } = string.Empty;
        public long SizeBytes { get; set; }
    }

    /// <summary>
    /// Nội dung ảnh đọc lại từ GridFS, dùng để FE hiển thị preview
    /// khi mở lại 1 ruleset đã có sẵn assetId.
    /// </summary>
    public class PictureBulletAssetContent
    {
        public byte[] Content { get; set; } = Array.Empty<byte>();
        public string ContentType { get; set; } = "application/octet-stream";
        public string FileName { get; set; } = string.Empty;
    }
}
