using MOS.ExcelGrading.Core.DTOs;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IGoogleSheetAttendanceSyncService
    {
        Task<GoogleSheetSyncResult?> SyncScheduleAttendanceAsync(
            ScheduleAttendanceResponse attendance,
            string requestedByUserId,
            bool throwOnError = false,
            CancellationToken cancellationToken = default);

        Task<GoogleSheetSyncResult?> SyncClassStudentMetadataAsync(
            string classId,
            string requestedByUserId,
            bool throwOnError = false,
            CancellationToken cancellationToken = default);
    }
}
