using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.API.Authorization;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Api.Controllers
{
    [ApiController]
    [Route("api/inserted-image-assets")]
    [Authorize(Roles = $"{UserRoles.Admin},{UserRoles.Teacher}")]
    public class InsertedImageAssetsController : ControllerBase
    {
        private readonly IImageAssetService _assetService;

        public InsertedImageAssetsController(IImageAssetService assetService)
        {
            _assetService = assetService;
        }

        // POST /api/inserted-image-assets
        // multipart/form-data, field "file"
        [HttpPost]
        [RequestSizeLimit(10 * 1024 * 1024)]
        [RequirePermission(Permissions.CreateXmlRules)]
        public async Task<IActionResult> Upload([FromForm] IFormFile? file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest(new { message = "Vui lòng chọn file ảnh." });
            }

            try
            {
                await using var stream = file.OpenReadStream();
                var result = await _assetService.UploadAsync(stream, file.FileName, file.ContentType, ImageAssetKind.InsertedImage);
                return Ok(result);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        // GET /api/inserted-image-assets/{assetId}
        [HttpGet("{assetId}")]
        [RequirePermission(Permissions.ViewXmlRules)]
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