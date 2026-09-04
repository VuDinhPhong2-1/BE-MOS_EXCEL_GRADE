namespace MOS.ExcelGrading.Core.Interfaces
{
    public enum ImageAssetKind
    {
        PictureBullet,
        InsertedImage
    }

    public interface IImageAssetService
    {
        Task<PictureBulletAssetUploadResult> UploadAsync(
            Stream content, string fileName, string contentType, ImageAssetKind kind);

        Task<PictureBulletAssetContent?> GetAsync(string assetId);
    }
}