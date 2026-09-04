using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models; 

namespace MOS.ExcelGrading.Api.Controllers
{
    [ApiController]
    [Route("api/picture-bullet-assets")]
    [Authorize(Roles = "Admin")]
    public class PictureBulletAssetsController : ControllerBase
    {
        private readonly IImageAssetService _assetService;

        public PictureBulletAssetsController(IImageAssetService assetService)
        {
            _assetService = assetService;
        }

        // POST /api/picture-bullet-assets
        // multipart/form-data, field "file"
        [HttpPost]
        [RequestSizeLimit(10 * 1024 * 1024)]
        public async Task<IActionResult> Upload([FromForm] IFormFile? file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest(new { message = "Vui lòng chọn file ảnh." });
            }

            try
            {
                await using var stream = file.OpenReadStream();
                var result = await _assetService.UploadAsync(stream, file.FileName, file.ContentType, ImageAssetKind.PictureBullet);
                return Ok(result);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        // GET /api/picture-bullet-assets/{assetId}
        [HttpGet("{assetId}")]
        public async Task<IActionResult> Get(string assetId)
        {
            var asset = await _assetService.GetAsync(assetId);
            if (asset == null)
            {
                return NotFound();
            }

            return File(asset.Content, asset.ContentType, asset.FileName);
        }
    }
}