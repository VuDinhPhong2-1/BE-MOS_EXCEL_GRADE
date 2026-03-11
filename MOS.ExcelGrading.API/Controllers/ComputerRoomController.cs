using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Security.Claims;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/computer-rooms")]
    [Authorize(Roles = $"{UserRoles.Admin},{UserRoles.Teacher}")]
    public class ComputerRoomController : ControllerBase
    {
        private readonly IComputerRoomService _computerRoomService;
        private readonly ISchoolService _schoolService;
        private readonly ILogger<ComputerRoomController> _logger;

        public ComputerRoomController(
            IComputerRoomService computerRoomService,
            ISchoolService schoolService,
            ILogger<ComputerRoomController> logger)
        {
            _computerRoomService = computerRoomService;
            _schoolService = schoolService;
            _logger = logger;
        }

        [HttpGet]
        public async Task<IActionResult> GetBySchool([FromQuery] string schoolId, [FromQuery] bool includeInactive = false)
        {
            if (string.IsNullOrWhiteSpace(schoolId))
            {
                return BadRequest(new { message = "Thiếu schoolId" });
            }

            var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
            var role = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(userId))
            {
                return Unauthorized(new { message = "Mã xác thực không hợp lệ" });
            }

            if (role != UserRoles.Admin)
            {
                var isOwner = await _schoolService.IsOwnerOfSchoolAsync(userId, schoolId);
                if (!isOwner)
                {
                    return Forbid();
                }
            }

            var rooms = await _computerRoomService.GetBySchoolIdAsync(schoolId, includeInactive);
            return Ok(rooms.Select(ToResponse).ToList());
        }

        [HttpPost]
        public async Task<IActionResult> Create([FromBody] CreateComputerRoomRequest request)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
            var role = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(userId))
            {
                return Unauthorized(new { message = "Mã xác thực không hợp lệ" });
            }

            var school = await _schoolService.GetSchoolByIdAsync(request.SchoolId.Trim());
            if (school == null || !school.IsActive)
            {
                return NotFound(new { message = "Không tìm thấy trường hoặc trường đã ngừng hoạt động" });
            }

            if (role != UserRoles.Admin && school.OwnerId != userId)
            {
                return Forbid();
            }

            var existed = await _computerRoomService.GetBySchoolAndNameAsync(school.Id!, request.Name.Trim(), true);
            if (existed != null)
            {
                return BadRequest(new { message = "Tên phòng máy đã tồn tại trong trường này" });
            }

            var room = new ComputerRoom
            {
                SchoolId = school.Id!,
                OwnerId = school.OwnerId,
                Name = request.Name.Trim(),
                StudentMachineCount = Math.Max(0, request.StudentMachineCount),
                TeacherMachineCount = Math.Max(0, request.TeacherMachineCount),
                BrokenMachineCount = Math.Max(0, request.BrokenMachineCount),
                NetSupportStatus = CleanOrDefault(request.NetSupportStatus, "Tốt"),
                AudioStatus = CleanOrDefault(request.AudioStatus, "Tốt"),
                CoolingStatus = CleanOrDefault(request.CoolingStatus, "Tốt"),
                DevicesPoweredOffStatus = CleanOrDefault(request.DevicesPoweredOffStatus, "Rồi"),
                SeatingOrderStatus = CleanOrDefault(request.SeatingOrderStatus, "Tốt"),
                RoomHygieneStatus = CleanOrDefault(request.RoomHygieneStatus, "Tốt"),
                IsActive = true
            };

            var created = await _computerRoomService.CreateAsync(room);
            return CreatedAtAction(nameof(GetById), new { id = created.Id }, ToResponse(created));
        }

        [HttpGet("{id}")]
        public async Task<IActionResult> GetById(string id)
        {
            var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
            var role = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(userId))
            {
                return Unauthorized(new { message = "Mã xác thực không hợp lệ" });
            }

            var room = await _computerRoomService.GetByIdAsync(id);
            if (room == null)
            {
                return NotFound(new { message = "Không tìm thấy phòng máy" });
            }

            if (role != UserRoles.Admin && room.OwnerId != userId)
            {
                return Forbid();
            }

            return Ok(ToResponse(room));
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> Update(string id, [FromBody] UpdateComputerRoomRequest request)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
            var role = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
            var isAdmin = role == UserRoles.Admin;
            if (string.IsNullOrWhiteSpace(userId))
            {
                return Unauthorized(new { message = "Mã xác thực không hợp lệ" });
            }

            var existing = await _computerRoomService.GetByIdAsync(id);
            if (existing == null)
            {
                return NotFound(new { message = "Không tìm thấy phòng máy" });
            }

            if (!isAdmin && existing.OwnerId != userId)
            {
                return Forbid();
            }

            var nextName = string.IsNullOrWhiteSpace(request.Name)
                ? existing.Name
                : request.Name.Trim();

            if (!string.Equals(nextName, existing.Name, StringComparison.OrdinalIgnoreCase))
            {
                var existed = await _computerRoomService.GetBySchoolAndNameAsync(existing.SchoolId, nextName, true);
                if (existed != null && !string.Equals(existed.Id, existing.Id, StringComparison.Ordinal))
                {
                    return BadRequest(new { message = "Tên phòng máy đã tồn tại trong trường này" });
                }
            }

            existing.Name = nextName;
            existing.StudentMachineCount = Math.Max(0, request.StudentMachineCount ?? existing.StudentMachineCount);
            existing.TeacherMachineCount = Math.Max(0, request.TeacherMachineCount ?? existing.TeacherMachineCount);
            existing.BrokenMachineCount = Math.Max(0, request.BrokenMachineCount ?? existing.BrokenMachineCount);
            existing.NetSupportStatus = CleanOrDefault(request.NetSupportStatus, existing.NetSupportStatus, "Tốt");
            existing.AudioStatus = CleanOrDefault(request.AudioStatus, existing.AudioStatus, "Tốt");
            existing.CoolingStatus = CleanOrDefault(request.CoolingStatus, existing.CoolingStatus, "Tốt");
            existing.DevicesPoweredOffStatus = CleanOrDefault(request.DevicesPoweredOffStatus, existing.DevicesPoweredOffStatus, "Rồi");
            existing.SeatingOrderStatus = CleanOrDefault(request.SeatingOrderStatus, existing.SeatingOrderStatus, "Tốt");
            existing.RoomHygieneStatus = CleanOrDefault(request.RoomHygieneStatus, existing.RoomHygieneStatus, "Tốt");
            existing.IsActive = request.IsActive ?? existing.IsActive;

            var updated = await _computerRoomService.UpdateAsync(id, existing, userId, isAdmin);
            if (updated == null)
            {
                _logger.LogWarning("Không thể cập nhật phòng máy {RoomId}", id);
                return NotFound(new { message = "Không thể cập nhật phòng máy" });
            }

            return Ok(ToResponse(updated));
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> Delete(string id)
        {
            var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
            var role = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
            var isAdmin = role == UserRoles.Admin;
            if (string.IsNullOrWhiteSpace(userId))
            {
                return Unauthorized(new { message = "Mã xác thực không hợp lệ" });
            }

            var deleted = await _computerRoomService.DeleteAsync(id, userId, isAdmin);
            if (!deleted)
            {
                return NotFound(new { message = "Không tìm thấy phòng máy hoặc không có quyền xóa" });
            }

            return Ok(new { message = "Đã xóa phòng máy" });
        }

        private static ComputerRoomResponse ToResponse(ComputerRoom room)
        {
            var availableStudentMachines = Math.Max(0, room.StudentMachineCount - room.BrokenMachineCount);
            return new ComputerRoomResponse
            {
                Id = room.Id ?? string.Empty,
                SchoolId = room.SchoolId,
                OwnerId = room.OwnerId,
                Name = room.Name,
                StudentMachineCount = room.StudentMachineCount,
                TeacherMachineCount = room.TeacherMachineCount,
                BrokenMachineCount = room.BrokenMachineCount,
                AvailableStudentMachines = availableStudentMachines,
                TotalMachineCount = room.StudentMachineCount + room.TeacherMachineCount,
                TotalMachinesText = $"{room.StudentMachineCount} + {room.TeacherMachineCount} GV",
                NetSupportStatus = room.NetSupportStatus,
                AudioStatus = room.AudioStatus,
                CoolingStatus = room.CoolingStatus,
                DevicesPoweredOffStatus = room.DevicesPoweredOffStatus,
                SeatingOrderStatus = room.SeatingOrderStatus,
                RoomHygieneStatus = room.RoomHygieneStatus,
                IsActive = room.IsActive,
                CreatedAt = room.CreatedAt,
                UpdatedAt = room.UpdatedAt
            };
        }

        private static string CleanOrDefault(string? value, string @default)
        {
            var cleaned = value?.Trim();
            return string.IsNullOrWhiteSpace(cleaned) ? @default : cleaned;
        }

        private static string CleanOrDefault(string? value, string? fallback, string @default)
        {
            var cleaned = value?.Trim();
            if (!string.IsNullOrWhiteSpace(cleaned))
            {
                return cleaned;
            }

            if (!string.IsNullOrWhiteSpace(fallback))
            {
                return fallback.Trim();
            }

            return @default;
        }
    }
}
