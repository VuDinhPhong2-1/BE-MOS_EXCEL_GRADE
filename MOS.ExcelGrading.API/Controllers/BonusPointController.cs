// MOS.ExcelGrading.API/Controllers/BonusPointController.cs
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.API.Authorization;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Security.Claims;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize(Roles = $"{UserRoles.Teacher},{UserRoles.Admin}")]
    public class BonusPointController : ControllerBase
    {
        private readonly IBonusPointService _bonusPointService;
        private readonly ILogger<BonusPointController> _logger;

        public BonusPointController(
            IBonusPointService bonusPointService,
            ILogger<BonusPointController> logger)
        {
            _bonusPointService = bonusPointService;
            _logger = logger;
        }

        /// <summary>
        /// Lấy tổng điểm cộng của toàn bộ học sinh trong lớp
        /// </summary>
        [HttpGet("class/{classId}")]
        [RequirePermission(Permissions.ViewBonusPoints)]
        public async Task<IActionResult> GetByClass(string classId)
        {
            try
            {
                var summary = await _bonusPointService.GetByClassAsync(classId);
                return Ok(summary);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting bonus points for class {ClassId}", classId);
                return StatusCode(500, new { message = ex.Message });
            }
        }

        /// <summary>
        /// Lấy điểm cộng của lớp theo ngày cụ thể (yyyy-MM-dd)
        /// </summary>
        [HttpGet("class/{classId}/date/{date}")]
        [RequirePermission(Permissions.ViewBonusPoints)]
        public async Task<IActionResult> GetByClassAndDate(string classId, DateTime date)
        {
            try
            {
                var points = await _bonusPointService.GetByClassAndDateAsync(classId, date);
                return Ok(points);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting bonus points for class {ClassId} on date {Date}", classId, date);
                return StatusCode(500, new { message = ex.Message });
            }
        }

        /// <summary>
        /// Lấy lịch sử điểm cộng của 1 học sinh trong lớp
        /// </summary>
        [HttpGet("student/{studentId}/class/{classId}")]
        [RequirePermission(Permissions.ViewBonusPoints)]
        public async Task<IActionResult> GetByStudent(string studentId, string classId)
        {
            try
            {
                var history = await _bonusPointService.GetByStudentAsync(studentId, classId);
                return Ok(history);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting bonus points for student {StudentId} in class {ClassId}", studentId, classId);
                return StatusCode(500, new { message = ex.Message });
            }
        }

        /// <summary>
        /// Tạo mới 1 bản ghi điểm cộng
        /// </summary>
        [HttpPost]
        [RequirePermission(Permissions.CreateBonusPoints)]
        public async Task<IActionResult> Create([FromBody] CreateBonusPointRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                var userId = GetUserId();
                var result = await _bonusPointService.CreateAsync(request, userId);
                return CreatedAtAction(nameof(GetByClass), new { classId = result.ClassId }, result);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating bonus point");
                return StatusCode(500, new { message = ex.Message });
            }
        }

        /// <summary>
        /// Nhập hàng loạt điểm cộng theo ngày (từ modal điểm danh hoặc trang quản lý)
        /// </summary>
        [HttpPost("bulk")]
        [RequirePermission(Permissions.CreateBonusPoints)]
        public async Task<IActionResult> BulkCreate([FromBody] BulkBonusPointRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                var userId = GetUserId();
                var results = await _bonusPointService.BulkCreateAsync(request, userId);
                return Ok(results);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error bulk creating bonus points for class {ClassId}", request.ClassId);
                return StatusCode(500, new { message = ex.Message });
            }
        }

        /// <summary>
        /// Cập nhật 1 bản ghi điểm cộng
        /// </summary>
        [HttpPut("{id}")]
        [RequirePermission(Permissions.EditBonusPoints)]
        public async Task<IActionResult> Update(string id, [FromBody] UpdateBonusPointRequest request)
        {
            try
            {
                var userId = GetUserId();
                var result = await _bonusPointService.UpdateAsync(id, request, userId);
                if (result == null)
                    return NotFound(new { message = "Không tìm thấy bản ghi điểm cộng." });

                return Ok(result);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating bonus point {Id}", id);
                return StatusCode(500, new { message = ex.Message });
            }
        }

        /// <summary>
        /// Xóa 1 bản ghi điểm cộng
        /// </summary>
        [HttpDelete("{id}")]
        [RequirePermission(Permissions.DeleteBonusPoints)]
        public async Task<IActionResult> Delete(string id)
        {
            try
            {
                var userId = GetUserId();
                var deleted = await _bonusPointService.DeleteAsync(id, userId);
                if (!deleted)
                    return NotFound(new { message = "Không tìm thấy bản ghi điểm cộng để xóa." });

                return Ok(new { message = "Xóa điểm cộng thành công." });
            }
            catch (ArgumentException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting bonus point {Id}", id);
                return StatusCode(500, new { message = ex.Message });
            }
        }

        private string GetUserId() =>
            User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
    }
}
