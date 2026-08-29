using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IPictureBulletAssetService
    {
        /// <summary>
        /// Upload ảnh bullet chuẩn, tính SHA-256, lưu vào GridFS.
        /// Ném InvalidOperationException nếu file không hợp lệ (sai định dạng, quá lớn, rỗng).
        /// </summary>
        Task<PictureBulletAssetUploadResult> UploadAsync(Stream content, string fileName, string contentType);

        /// <summary>
        /// Đọc lại ảnh theo assetId để hiển thị preview. Trả về null nếu không tìm thấy.
        /// </summary>
        Task<PictureBulletAssetContent?> GetAsync(string assetId);
    }
}
