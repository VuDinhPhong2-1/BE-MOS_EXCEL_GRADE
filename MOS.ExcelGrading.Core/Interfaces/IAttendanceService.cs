using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IAttendanceService
    {
        Task<ScheduleAttendanceResponse> GetScheduleAttendanceAsync(TeacherSchedule schedule, string ownerId, bool isAdmin);
        Task<ScheduleAttendanceResponse> SaveScheduleAttendanceAsync(
            TeacherSchedule schedule,
            string ownerId,
            bool isAdmin,
            IReadOnlyCollection<SaveScheduleAttendanceItem> items,
            ScheduleReportsRequest? reports,
            string updatedBy);
    }
}
