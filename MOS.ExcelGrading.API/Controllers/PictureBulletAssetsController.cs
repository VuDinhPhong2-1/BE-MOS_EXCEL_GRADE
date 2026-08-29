using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models; 

namespace MOS.ExcelGrading.Api.Controllers
{
    // Lưu ý: điều chỉnh namespace/route-prefix/attribute Authorize cho khớp với
    // cách các controller khác trong project đang được cấu hình (ví dụ nếu có
    // [Authorize(Policy = "AdminOnly")] riêng thì dùng lại policy đó thay vì Roles).
    [ApiController]
    [Route("api/picture-bullet-assets")]
    [Authorize(Roles = "Admin")]
    public class PictureBulletAssetsController : ControllerBase
    {
        private readonly IPictureBulletAssetService _assetService;

        public PictureBulletAssetsController(IPictureBulletAssetService assetService)
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
                var result = await _assetService.UploadAsync(stream, file.FileName, file.ContentType);
                return Ok(result);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        // GET /api/picture-bullet-assets/{assetId}
        // Dùng để FE hiển thị lại preview khi mở 1 ruleset đã có sẵn assetId.
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
