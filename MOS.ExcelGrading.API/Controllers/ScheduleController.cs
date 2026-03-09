using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Security.Claims;

namespace MOS.ExcelGrading.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class ScheduleController : ControllerBase
    {
        private static readonly TimeZoneInfo VietnamTimeZone = ResolveVietnamTimeZone();

        private readonly IScheduleService _scheduleService;
        private readonly IClassService _classService;
        private readonly IAttendanceService _attendanceService;
        private readonly ILogger<ScheduleController> _logger;

        public ScheduleController(
            IScheduleService scheduleService,
            IClassService classService,
            IAttendanceService attendanceService,
            ILogger<ScheduleController> logger)
        {
            _scheduleService = scheduleService;
            _classService = classService;
            _attendanceService = attendanceService;
            _logger = logger;
        }

        [HttpGet("week")]
        public async Task<IActionResult> GetWeekSchedules([FromQuery] DateTime? weekStart, [FromQuery] bool includeInactive = false)
        {
            try
            {
                var ownerId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                if (string.IsNullOrWhiteSpace(ownerId))
                    return Unauthorized(new { message = "Mã xác thực không hợp lệ" });

                var localBaseDate = weekStart?.Date ?? GetVietnamToday();
                var localWeekStart = GetWeekStartLocal(localBaseDate);
                var utcWeekStart = ToUtcFromVietnamLocalDate(localWeekStart);

                var items = await _scheduleService.GetSchedulesByWeekAsync(ownerId, utcWeekStart, includeInactive);

                var response = items.Select(ToResponse).ToList();
                return Ok(new
                {
                    weekStart = localWeekStart,
                    weekEnd = localWeekStart.AddDays(6),
                    data = response
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi lấy lịch dạy theo tuần");
                return StatusCode(500, new { message = "Đã xảy ra lỗi khi lấy lịch dạy" });
            }
        }

        [HttpGet("{id}")]
        public async Task<IActionResult> GetById(string id)
        {
            try
            {
                var ownerId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var role = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
                var isAdmin = role == UserRoles.Admin;

                var schedule = await _scheduleService.GetByIdAsync(id);
                if (schedule == null)
                    return NotFound(new { message = "Không tìm thấy lịch dạy" });

                if (!isAdmin && schedule.OwnerId != ownerId)
                    return Forbid();

                return Ok(ToResponse(schedule));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi lấy chi tiết lịch dạy");
                return StatusCode(500, new { message = "Đã xảy ra lỗi khi lấy lịch dạy" });
            }
        }

        [HttpPost]
        public async Task<IActionResult> Create([FromBody] CreateScheduleRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                if (!IsTimeRangeValid(request.StartTime, request.EndTime))
                    return BadRequest(new { message = "Giờ kết thúc phải lớn hơn giờ bắt đầu" });

                var ownerId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var role = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
                var isAdmin = role == UserRoles.Admin;
                if (string.IsNullOrWhiteSpace(ownerId))
                    return Unauthorized(new { message = "Mã xác thực không hợp lệ" });

                var classValidation = await ValidateAndResolveClassContextAsync(
                    ownerId,
                    isAdmin,
                    request.ClassId,
                    request.SchoolId);
                if (!classValidation.IsValid)
                {
                    return classValidation.ErrorResult!;
                }

                var schedule = new TeacherSchedule
                {
                    OwnerId = ownerId,
                    SchoolId = classValidation.SchoolId,
                    ClassId = classValidation.ClassId,
                    ClassName = classValidation.ClassName,
                    Subject = request.Subject.Trim(),
                    RoomName = request.RoomName?.Trim(),
                    PeriodLabel = request.PeriodLabel?.Trim(),
                    Date = ToUtcFromVietnamLocalDate(request.Date.Date),
                    StartTime = request.StartTime.Trim(),
                    EndTime = request.EndTime.Trim(),
                    Notes = request.Notes?.Trim(),
                    IsActive = true
                };

                var created = await _scheduleService.CreateAsync(schedule);
                return CreatedAtAction(nameof(GetById), new { id = created.Id }, ToResponse(created));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi tạo lịch dạy");
                return StatusCode(500, new { message = "Đã xảy ra lỗi khi tạo lịch dạy" });
            }
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> Update(string id, [FromBody] UpdateScheduleRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                if (!IsTimeRangeValid(request.StartTime, request.EndTime))
                    return BadRequest(new { message = "Giờ kết thúc phải lớn hơn giờ bắt đầu" });

                var ownerId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var role = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
                var isAdmin = role == UserRoles.Admin;
                if (string.IsNullOrWhiteSpace(ownerId))
                    return Unauthorized(new { message = "Mã xác thực không hợp lệ" });

                var existing = await _scheduleService.GetByIdAsync(id);
                if (existing == null)
                    return NotFound(new { message = "Không tìm thấy lịch dạy" });

                var canAccess = isAdmin
                    || existing.OwnerId == ownerId
                    || string.IsNullOrWhiteSpace(existing.OwnerId);

                if (!canAccess)
                    return Forbid();

                var classValidation = await ValidateAndResolveClassContextAsync(
                    ownerId,
                    isAdmin,
                    request.ClassId,
                    request.SchoolId);
                if (!classValidation.IsValid)
                {
                    return classValidation.ErrorResult!;
                }

                var schedule = new TeacherSchedule
                {
                    OwnerId = string.IsNullOrWhiteSpace(existing.OwnerId) ? ownerId : existing.OwnerId,
                    SchoolId = classValidation.SchoolId,
                    ClassId = classValidation.ClassId,
                    ClassName = classValidation.ClassName,
                    Subject = request.Subject.Trim(),
                    RoomName = request.RoomName?.Trim(),
                    PeriodLabel = request.PeriodLabel?.Trim(),
                    Date = ToUtcFromVietnamLocalDate(request.Date.Date),
                    StartTime = request.StartTime.Trim(),
                    EndTime = request.EndTime.Trim(),
                    Notes = request.Notes?.Trim(),
                    IsActive = request.IsActive ?? true
                };

                var updated = await _scheduleService.UpdateAsync(id, schedule, ownerId, true);
                if (updated == null)
                    return NotFound(new { message = "Không tìm thấy lịch dạy hoặc bạn không có quyền sửa" });

                return Ok(ToResponse(updated));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi cập nhật lịch dạy");
                return StatusCode(500, new { message = "Đã xảy ra lỗi khi cập nhật lịch dạy" });
            }
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> Delete(string id)
        {
            try
            {
                var ownerId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var role = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
                var isAdmin = role == UserRoles.Admin;
                if (string.IsNullOrWhiteSpace(ownerId))
                    return Unauthorized(new { message = "Mã xác thực không hợp lệ" });

                var existing = await _scheduleService.GetByIdAsync(id);
                if (existing == null)
                    return NotFound(new { message = "Không tìm thấy lịch dạy" });

                var canAccess = isAdmin
                    || existing.OwnerId == ownerId
                    || string.IsNullOrWhiteSpace(existing.OwnerId);
                if (!canAccess)
                    return Forbid();

                var deleted = await _scheduleService.DeleteAsync(id, ownerId, true);
                if (!deleted)
                    return NotFound(new { message = "Không tìm thấy lịch dạy hoặc bạn không có quyền xóa" });

                return Ok(new { message = "Đã xóa lịch dạy" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi xóa lịch dạy");
                return StatusCode(500, new { message = "Đã xảy ra lỗi khi xóa lịch dạy" });
            }
        }

        [HttpGet("{id}/attendance")]
        public async Task<IActionResult> GetAttendance(string id)
        {
            try
            {
                var ownerId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var role = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
                var isAdmin = role == UserRoles.Admin;
                if (string.IsNullOrWhiteSpace(ownerId))
                    return Unauthorized(new { message = "Mã xác thực không hợp lệ" });

                var existing = await _scheduleService.GetByIdAsync(id);
                if (existing == null)
                    return NotFound(new { message = "Không tìm thấy lịch dạy" });

                var canAccess = isAdmin
                    || existing.OwnerId == ownerId
                    || string.IsNullOrWhiteSpace(existing.OwnerId);

                if (!canAccess)
                    return Forbid();

                var response = await _attendanceService.GetScheduleAttendanceAsync(existing, ownerId, isAdmin);
                response.Date = ToVietnamLocalDate(response.Date);
                return Ok(response);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi lấy điểm danh theo lịch");
                return StatusCode(500, new { message = "Đã xảy ra lỗi khi lấy điểm danh" });
            }
        }

        [HttpPut("{id}/attendance")]
        public async Task<IActionResult> SaveAttendance(string id, [FromBody] SaveScheduleAttendanceRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                    return BadRequest(ModelState);

                var ownerId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
                var role = User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
                var isAdmin = role == UserRoles.Admin;
                if (string.IsNullOrWhiteSpace(ownerId))
                    return Unauthorized(new { message = "Mã xác thực không hợp lệ" });

                var existing = await _scheduleService.GetByIdAsync(id);
                if (existing == null)
                    return NotFound(new { message = "Không tìm thấy lịch dạy" });

                var canAccess = isAdmin
                    || existing.OwnerId == ownerId
                    || string.IsNullOrWhiteSpace(existing.OwnerId);

                if (!canAccess)
                    return Forbid();

                var result = await _attendanceService.SaveScheduleAttendanceAsync(
                    existing,
                    ownerId,
                    isAdmin,
                    request.Items,
                    ownerId);

                result.Date = ToVietnamLocalDate(result.Date);
                return Ok(result);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi lưu điểm danh theo lịch");
                return StatusCode(500, new { message = "Đã xảy ra lỗi khi lưu điểm danh" });
            }
        }

        private async Task<ClassContextValidationResult> ValidateAndResolveClassContextAsync(
            string ownerId,
            bool isAdmin,
            string? classId,
            string? schoolIdFromRequest)
        {
            if (string.IsNullOrWhiteSpace(classId))
            {
                return ClassContextValidationResult.Fail(
                    BadRequest(new { message = "Vui lòng chọn lớp để tạo/cập nhật lịch dạy" }));
            }

            var classEntity = await _classService.GetClassByIdAsync(classId.Trim());
            if (classEntity == null || !classEntity.IsActive)
            {
                return ClassContextValidationResult.Fail(
                    NotFound(new { message = "Không tìm thấy lớp hoặc lớp đã ngừng hoạt động" }));
            }

            if (!isAdmin && classEntity.OwnerId != ownerId)
            {
                return ClassContextValidationResult.Fail(Forbid());
            }

            if (!string.IsNullOrWhiteSpace(schoolIdFromRequest) &&
                !string.Equals(classEntity.SchoolId, schoolIdFromRequest.Trim(), StringComparison.Ordinal))
            {
                return ClassContextValidationResult.Fail(
                    BadRequest(new { message = "Lớp đã chọn không thuộc trường đã chọn" }));
            }

            return ClassContextValidationResult.Success(
                classEntity.Id ?? classId.Trim(),
                classEntity.Name,
                classEntity.SchoolId);
        }

        private static DateTime GetVietnamToday()
        {
            return TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, VietnamTimeZone).Date;
        }

        private static DateTime GetWeekStartLocal(DateTime localDate)
        {
            var d = localDate.Date;
            var diff = (7 + (int)d.DayOfWeek - (int)DayOfWeek.Monday) % 7;
            return d.AddDays(-diff);
        }

        private static DateTime ToUtcFromVietnamLocalDate(DateTime localDate)
        {
            var unspecified = new DateTime(localDate.Year, localDate.Month, localDate.Day, 0, 0, 0, DateTimeKind.Unspecified);
            return TimeZoneInfo.ConvertTimeToUtc(unspecified, VietnamTimeZone);
        }

        private static DateTime ToVietnamLocalDate(DateTime value)
        {
            var utc = value.Kind switch
            {
                DateTimeKind.Utc => value,
                DateTimeKind.Local => value.ToUniversalTime(),
                _ => DateTime.SpecifyKind(value, DateTimeKind.Utc)
            };

            return TimeZoneInfo.ConvertTimeFromUtc(utc, VietnamTimeZone).Date;
        }

        private static TimeZoneInfo ResolveVietnamTimeZone()
        {
            try
            {
                return TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time");
            }
            catch
            {
                try
                {
                    return TimeZoneInfo.FindSystemTimeZoneById("Asia/Ho_Chi_Minh");
                }
                catch
                {
                    return TimeZoneInfo.Utc;
                }
            }
        }

        private static bool IsTimeRangeValid(string startTime, string endTime)
        {
            if (!TimeSpan.TryParse(startTime, out var start)) return false;
            if (!TimeSpan.TryParse(endTime, out var end)) return false;
            return end > start;
        }

        private static ScheduleResponse ToResponse(TeacherSchedule schedule)
        {
            var localDate = ToVietnamLocalDate(schedule.Date);
            return new ScheduleResponse
            {
                Id = schedule.Id ?? string.Empty,
                OwnerId = schedule.OwnerId,
                SchoolId = schedule.SchoolId,
                ClassId = schedule.ClassId,
                ClassName = schedule.ClassName,
                Subject = schedule.Subject,
                RoomName = schedule.RoomName,
                PeriodLabel = schedule.PeriodLabel,
                Date = localDate,
                DayOfWeek = ConvertToVietnameseDayOfWeek(localDate.DayOfWeek),
                StartTime = schedule.StartTime,
                EndTime = schedule.EndTime,
                Notes = schedule.Notes,
                IsActive = schedule.IsActive,
                CreatedAt = schedule.CreatedAt,
                UpdatedAt = schedule.UpdatedAt
            };
        }

        private static int ConvertToVietnameseDayOfWeek(DayOfWeek dayOfWeek)
        {
            return dayOfWeek switch
            {
                DayOfWeek.Monday => 2,
                DayOfWeek.Tuesday => 3,
                DayOfWeek.Wednesday => 4,
                DayOfWeek.Thursday => 5,
                DayOfWeek.Friday => 6,
                DayOfWeek.Saturday => 7,
                _ => 8
            };
        }

        private readonly record struct ClassContextValidationResult(
            bool IsValid,
            IActionResult? ErrorResult,
            string ClassId,
            string ClassName,
            string SchoolId)
        {
            public static ClassContextValidationResult Fail(IActionResult errorResult)
                => new(false, errorResult, string.Empty, string.Empty, string.Empty);

            public static ClassContextValidationResult Success(string classId, string className, string schoolId)
                => new(true, null, classId, className, schoolId);
        }
    }
}
