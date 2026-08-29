using System.Security.Cryptography;

namespace MOS.ExcelGrading.Core.Utilities
{
    /// <summary>
    /// Hàm hash dùng chung cho việc so khớp ảnh (Picture Bullet, và các tính năng
    /// tương lai cần so ảnh bằng SHA-256). Tách ra từ XmlGradingRuleService để
    /// PictureBulletAssetService cũng dùng chung logic, tránh 2 nơi tính hash
    /// theo 2 cách khác nhau dẫn tới không khớp.
    /// </summary>
    public static class ImageHashUtility
    {
        public static string ComputeSha256(byte[] bytes)
        {
            using var sha256 = SHA256.Create();
            var hash = sha256.ComputeHash(bytes);
            return Convert.ToHexString(hash).ToLowerInvariant();
        }

        public static string NormalizeHash(string hash)
        {
            return (hash ?? string.Empty).Trim().Replace("-", "").Replace(" ", "").ToLowerInvariant();
        }
    }
}
