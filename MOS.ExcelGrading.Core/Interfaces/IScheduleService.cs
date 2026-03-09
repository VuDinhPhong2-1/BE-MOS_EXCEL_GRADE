using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IScheduleService
    {
        Task<List<TeacherSchedule>> GetSchedulesByWeekAsync(string ownerId, DateTime weekStart, bool includeInactive = false);
        Task<TeacherSchedule?> GetByIdAsync(string id);
        Task<TeacherSchedule> CreateAsync(TeacherSchedule schedule);
        Task<TeacherSchedule?> UpdateAsync(string id, TeacherSchedule schedule, string updatedBy, bool isAdmin);
        Task<bool> DeleteAsync(string id, string ownerId, bool isAdmin);
    }
}
