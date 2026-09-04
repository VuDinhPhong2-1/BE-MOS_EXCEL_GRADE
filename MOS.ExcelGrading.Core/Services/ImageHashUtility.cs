using System.Security.Cryptography;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
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

        public static string ComputePerceptualHash(byte[] imageBytes)
    {
        using var image = Image.Load<L8>(imageBytes); // L8 = grayscale 8-bit
        image.Mutate(ctx => ctx.Resize(9, 8)); // 9 cột để lấy 8 hiệu số/hàng

        ulong hash = 0;
        int bitIndex = 0;

        for (int y = 0; y < 8; y++)
        {
            for (int x = 0; x < 8; x++)
            {
                var left = image[x, y].PackedValue;
                var right = image[x + 1, y].PackedValue;

                if (left < right)
                {
                    hash |= (1UL << bitIndex);
                }

                bitIndex++;
            }
        }

        return hash.ToString("x16");
    }

    /// <summary>
    /// Số bit khác nhau giữa 2 perceptual hash (Hamming distance).
    /// 0 = giống hệt, càng lớn càng khác biệt. Với dHash 64-bit,
    /// ngưỡng thường dùng: <= 8-10 coi là "cùng ảnh" sau khi nén lại,
    /// > 20 gần như chắc chắn là ảnh khác.
    /// </summary>
    public static int HammingDistance(string hashA, string hashB)
    {
        var a = Convert.ToUInt64(hashA, 16);
        var b = Convert.ToUInt64(hashB, 16);
        var xor = a ^ b;
        return System.Numerics.BitOperations.PopCount(xor);
    }

    public static bool IsPerceptuallySimilar(string hashA, string hashB, int threshold = 10)
    {
        return HammingDistance(hashA, hashB) <= threshold;
    }
    }
    
}
