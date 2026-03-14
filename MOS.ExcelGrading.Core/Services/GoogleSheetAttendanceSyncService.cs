using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class GoogleSheetAttendanceSyncService : IGoogleSheetAttendanceSyncService
    {
        private const int AttendanceStartColumnIndex = 9; // I
        private const int StudentStartRow = 8;
        private const string PresenceMark = "x";

        private readonly IClassService _classService;
        private readonly ISchoolService _schoolService;
        private readonly IStudentService _studentService;
        private readonly IOptions<GoogleSheetsSettings> _settings;
        private readonly ILogger<GoogleSheetAttendanceSyncService> _logger;

        public GoogleSheetAttendanceSyncService(
            IClassService classService,
            ISchoolService schoolService,
            IStudentService studentService,
            IOptions<GoogleSheetsSettings> settings,
            ILogger<GoogleSheetAttendanceSyncService> logger)
        {
            _classService = classService;
            _schoolService = schoolService;
            _studentService = studentService;
            _settings = settings;
            _logger = logger;
        }

        public async Task<GoogleSheetSyncResult?> SyncScheduleAttendanceAsync(
            ScheduleAttendanceResponse attendance,
            string requestedByUserId,
            bool throwOnError = false,
            CancellationToken cancellationToken = default)
        {
            // Kiểm tra đầu vào: bắt buộc phải có ClassId để tìm cấu hình sheet theo từng lớp.
            if (attendance == null || string.IsNullOrWhiteSpace(attendance.ClassId))
            {
                if (throwOnError)
                {
                    throw new InvalidOperationException("Thiếu thông tin lớp học để đồng bộ Google Sheet.");
                }
                return null;
            }

            // Đọc cấu hình Google Sheets toàn hệ thống (Enabled, service account, default spreadsheet).
            var settings = _settings.Value;
            // Nếu tính năng Google Sheets đang tắt thì dừng tại đây (hoặc throw nếu throwOnError=true).
            if (!settings.Enabled)
            {
                if (throwOnError)
                {
                    throw new InvalidOperationException("Google Sheets chưa được bật cấu hình (GoogleSheets:Enabled=false).");
                }
                return null;
            }

            try
            {
                // Tìm class trước để lấy cấu hình spreadsheet/tab riêng của lớp.
                var classEntity = await _classService.GetClassByIdAsync(attendance.ClassId);
                if (classEntity == null)
                {
                    if (throwOnError)
                    {
                        throw new InvalidOperationException($"Không tìm thấy lớp {attendance.ClassId} để đồng bộ Google Sheet.");
                    }

                    _logger.LogWarning(
                        "Skip Google Sheet sync: class not found. ClassId={ClassId}, ScheduleId={ScheduleId}",
                        attendance.ClassId,
                        attendance.ScheduleId);
                    return null;
                }

                // SpreadsheetId lấy duy nhất theo cấu hình của Trường: School.AttendanceSpreadsheetId.
                var schoolEntity = string.IsNullOrWhiteSpace(classEntity.SchoolId)
                    ? null
                    : await _schoolService.GetSchoolByIdAsync(classEntity.SchoolId);

                var spreadsheetId = ParseSpreadsheetId(schoolEntity?.AttendanceSpreadsheetId) ?? string.Empty;
                if (string.IsNullOrWhiteSpace(spreadsheetId))
                {
                    if (throwOnError)
                    {
                        throw new InvalidOperationException(
                            "Chưa cấu hình AttendanceSpreadsheetId cho trường.");
                    }

                    _logger.LogInformation(
                        "Skip Google Sheet sync: missing spreadsheet id/url. ClassId={ClassId}, ClassName={ClassName}",
                        classEntity.Id,
                        classEntity.Name);
                    return null;
                }

                // Thứ tự ưu tiên tên tab: class.AttendanceWorksheetName -> attendance.ClassName -> class.Name.
                var worksheetName = FirstNonEmpty(
                    classEntity.AttendanceWorksheetName,
                    attendance.ClassName,
                    classEntity.Name);
                if (string.IsNullOrWhiteSpace(worksheetName))
                {
                    if (throwOnError)
                    {
                        throw new InvalidOperationException(
                            "Chưa cấu hình AttendanceWorksheetName cho lớp và không suy ra được tên tab từ lịch dạy.");
                    }

                    _logger.LogInformation(
                        "Skip Google Sheet sync: missing worksheet name. ClassId={ClassId}, SpreadsheetId={SpreadsheetId}",
                        classEntity.Id,
                        spreadsheetId);
                    return null;
                }

                // Khởi tạo Google Sheets client từ service-account credentials.
                var sheetsService = BuildSheetsService(settings);
                if (sheetsService == null)
                {
                    if (throwOnError)
                    {
                        throw new InvalidOperationException(
                            "Thiếu cấu hình Service Account Google Sheets (ServiceAccountJson hoặc ServiceAccountJsonPath).");
                    }

                    _logger.LogWarning("Skip Google Sheet sync: service account is not configured.");
                    return null;
                }

                // Thực hiện đồng bộ dữ liệu điểm danh lên Google Sheet.
                var syncResult = await SyncToSheetAsync(
                    sheetsService,
                    spreadsheetId,
                    worksheetName,
                    attendance,
                    cancellationToken);

                _logger.LogInformation(
                    "Google Sheet sync succeeded. ScheduleId={ScheduleId}, SpreadsheetId={SpreadsheetId}, Worksheet={Worksheet}, Column={Column}, Matched={Matched}/{Total}, UserId={UserId}",
                    attendance.ScheduleId,
                    spreadsheetId,
                    worksheetName,
                    syncResult.ColumnLetter,
                    syncResult.MatchedStudentCount,
                    syncResult.TotalStudentCount,
                    requestedByUserId);

                return syncResult;
            }
            catch (Exception ex)
            {
                // Log cảnh báo để theo dõi lỗi đồng bộ, không làm mất dữ liệu điểm danh trong DB.
                _logger.LogWarning(
                    ex,
                    "Google Sheet sync failed. ScheduleId={ScheduleId}, ClassId={ClassId}",
                    attendance.ScheduleId,
                    attendance.ClassId);

                if (throwOnError)
                {
                    // Lỗi nghiệp vụ đã rõ ràng thì throw lại nguyên bản.
                    if (ex is InvalidOperationException)
                    {
                        throw;
                    }

                    // Nếu là lỗi khác (Google API, tab sai, quyền sai...) thì bổ sung context để FE debug.
                    var classConfig = await _classService.GetClassByIdAsync(attendance.ClassId);
                    var schoolConfig = classConfig == null || string.IsNullOrWhiteSpace(classConfig.SchoolId)
                        ? null
                        : await _schoolService.GetSchoolByIdAsync(classConfig.SchoolId);
                    var effectiveSpreadsheetId = ParseSpreadsheetId(schoolConfig?.AttendanceSpreadsheetId) ?? string.Empty;
                    var effectiveWorksheet = FirstNonEmpty(
                        classConfig?.AttendanceWorksheetName,
                        attendance.ClassName,
                        classConfig?.Name);
                    var detail = BuildSyncErrorDetail(ex, effectiveSpreadsheetId, effectiveWorksheet);

                    throw new InvalidOperationException(
                        $"Đồng bộ Google Sheet thất bại: {detail}",
                        ex);
                }
                return null;
            }
        }

        public async Task<GoogleSheetSyncResult?> SyncClassStudentMetadataAsync(
            string classId,
            string requestedByUserId,
            bool throwOnError = false,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(classId))
            {
                if (throwOnError)
                {
                    throw new InvalidOperationException("Thiếu thông tin lớp học để đồng bộ Google Sheet.");
                }
                return null;
            }

            var settings = _settings.Value;
            if (!settings.Enabled)
            {
                if (throwOnError)
                {
                    throw new InvalidOperationException("Google Sheets chưa được bật cấu hình (GoogleSheets:Enabled=false).");
                }
                return null;
            }

            try
            {
                var classEntity = await _classService.GetClassByIdAsync(classId);
                if (classEntity == null)
                {
                    if (throwOnError)
                    {
                        throw new InvalidOperationException($"Không tìm thấy lớp {classId} để đồng bộ Google Sheet.");
                    }

                    _logger.LogWarning(
                        "Skip Google Sheet metadata sync: class not found. ClassId={ClassId}",
                        classId);
                    return null;
                }

                var schoolEntity = string.IsNullOrWhiteSpace(classEntity.SchoolId)
                    ? null
                    : await _schoolService.GetSchoolByIdAsync(classEntity.SchoolId);

                var spreadsheetId = ParseSpreadsheetId(schoolEntity?.AttendanceSpreadsheetId) ?? string.Empty;
                if (string.IsNullOrWhiteSpace(spreadsheetId))
                {
                    if (throwOnError)
                    {
                        throw new InvalidOperationException(
                            "Chưa cấu hình AttendanceSpreadsheetId cho trường.");
                    }

                    _logger.LogInformation(
                        "Skip Google Sheet metadata sync: missing spreadsheet id/url. ClassId={ClassId}, ClassName={ClassName}",
                        classEntity.Id,
                        classEntity.Name);
                    return null;
                }

                var worksheetName = FirstNonEmpty(
                    classEntity.AttendanceWorksheetName,
                    classEntity.Name);
                if (string.IsNullOrWhiteSpace(worksheetName))
                {
                    if (throwOnError)
                    {
                        throw new InvalidOperationException(
                            "Chưa cấu hình AttendanceWorksheetName cho lớp và không suy ra được tên tab.");
                    }

                    _logger.LogInformation(
                        "Skip Google Sheet metadata sync: missing worksheet name. ClassId={ClassId}, SpreadsheetId={SpreadsheetId}",
                        classEntity.Id,
                        spreadsheetId);
                    return null;
                }

                var sheetsService = BuildSheetsService(settings);
                if (sheetsService == null)
                {
                    if (throwOnError)
                    {
                        throw new InvalidOperationException(
                            "Thiếu cấu hình Service Account Google Sheets (ServiceAccountJson hoặc ServiceAccountJsonPath).");
                    }

                    _logger.LogWarning("Skip Google Sheet metadata sync: service account is not configured.");
                    return null;
                }

                var students = await _studentService.GetByClassIdAsync(classId);
                var syncResult = await SyncStudentMetadataToSheetAsync(
                    sheetsService,
                    spreadsheetId,
                    worksheetName,
                    students,
                    cancellationToken);

                _logger.LogInformation(
                    "Google Sheet metadata sync succeeded. ClassId={ClassId}, SpreadsheetId={SpreadsheetId}, Worksheet={Worksheet}, ClassificationCol={ClassificationCol}, NotesCol={NotesCol}, Matched={Matched}/{Total}, UserId={UserId}",
                    classId,
                    spreadsheetId,
                    worksheetName,
                    syncResult.ClassificationColumnLetter,
                    syncResult.NotesColumnLetter,
                    syncResult.MatchedStudentCount,
                    syncResult.TotalStudentCount,
                    requestedByUserId);

                return syncResult;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(
                    ex,
                    "Google Sheet metadata sync failed. ClassId={ClassId}",
                    classId);

                if (throwOnError)
                {
                    if (ex is InvalidOperationException)
                    {
                        throw;
                    }

                    var classConfig = await _classService.GetClassByIdAsync(classId);
                    var schoolConfig = classConfig == null || string.IsNullOrWhiteSpace(classConfig.SchoolId)
                        ? null
                        : await _schoolService.GetSchoolByIdAsync(classConfig.SchoolId);
                    var effectiveSpreadsheetId = ParseSpreadsheetId(schoolConfig?.AttendanceSpreadsheetId) ?? string.Empty;
                    var effectiveWorksheet = FirstNonEmpty(
                        classConfig?.AttendanceWorksheetName,
                        classConfig?.Name);
                    var detail = BuildSyncErrorDetail(ex, effectiveSpreadsheetId, effectiveWorksheet);

                    throw new InvalidOperationException(
                        $"Đồng bộ metadata học sinh lên Google Sheet thất bại: {detail}",
                        ex);
                }

                return null;
            }
        }

        private static string BuildSyncErrorDetail(Exception ex, string spreadsheetId, string worksheetName)
        {
            if (ex is Google.GoogleApiException googleEx)
            {
                if (googleEx.HttpStatusCode == HttpStatusCode.Forbidden
                    || googleEx.HttpStatusCode == HttpStatusCode.Unauthorized)
                {
                    return "Service account chưa có quyền Editor trên file Google Sheet.";
                }

                if (googleEx.HttpStatusCode == HttpStatusCode.NotFound)
                {
                    return $"Không tìm thấy spreadsheet hoặc tab. SpreadsheetId='{spreadsheetId}', Tab='{worksheetName}'.";
                }

                var googleMessage = googleEx.Error?.Message;
                if (!string.IsNullOrWhiteSpace(googleMessage))
                {
                    return $"{googleMessage} (SpreadsheetId='{spreadsheetId}', Tab='{worksheetName}')";
                }
            }

            if (ex.Message.Contains("Unable to parse range", StringComparison.OrdinalIgnoreCase))
            {
                return $"Tên tab không tồn tại hoặc sai ký tự. Tab hiện tại: '{worksheetName}'.";
            }

            if (ex.Message.Contains("Requested entity was not found", StringComparison.OrdinalIgnoreCase))
            {
                return $"SpreadsheetId hoặc tab không đúng. SpreadsheetId='{spreadsheetId}', Tab='{worksheetName}'.";
            }

            return $"{ex.Message} (SpreadsheetId='{spreadsheetId}', Tab='{worksheetName}')";
        }

        private async Task<GoogleSheetSyncResult> SyncToSheetAsync(
            SheetsService sheetsService,
            string spreadsheetId,
            string worksheetName,
            ScheduleAttendanceResponse attendance,
            CancellationToken cancellationToken)
        {
            // Tìm cột buổi học cần ghi (cột I trở đi), dựa theo ngày/slot.
            var targetColumnIndex = await ResolveTargetColumnIndexAsync(
                sheetsService,
                spreadsheetId,
                worksheetName,
                attendance,
                cancellationToken);

            // Chuyển index cột (ví dụ 9) thành chữ cái cột (ví dụ I).
            var targetColumnLetter = ToColumnLetter(targetColumnIndex);

            // Khởi tạo danh sách update sẽ gửi 1 lần bằng BatchUpdate.
            var updates = new List<ValueRange>
            {
                // Không cập nhật hàng 5 (số buổi/Bxx), vì đây là cột template cố định của sheet.
                // Hàng 6: số tiết của buổi học.
                BuildSingleCellUpdate(
                    BuildRange(worksheetName, $"{targetColumnLetter}6"),
                    ResolvePeriodCount(attendance)),

                // Hàng 7: ngày dạy theo định dạng dd/MM/yyyy.
                BuildSingleCellUpdate(
                    BuildRange(worksheetName, $"{targetColumnLetter}7"),
                    attendance.Date.Date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture))
            };

            // Đọc các dòng học sinh trên sheet (B:C, từ dòng 8 trở đi).
            var sheetRows = await LoadSheetStudentRowsAsync(
                sheetsService,
                spreadsheetId,
                worksheetName,
                cancellationToken);

            // Tạo lookup map tên kiểu tách họ|tên.
            var studentsBySplitName = BuildStudentLookup(attendance.Students, useFullName: false);
            // Tạo lookup map tên kiểu full name (fallback khi không khớp kiểu trên).
            var studentsByFullName = BuildStudentLookup(attendance.Students, useFullName: true);
            // Lưu danh sách StudentId đã được map để tránh map trùng.
            var consumedStudentIds = new HashSet<string>(StringComparer.Ordinal);

            // Số lượng học sinh map được với dòng sheet.
            var matchedStudentCount = 0;
            // Số lượng học sinh được đánh dấu "có mặt" (ghi x).
            var presentMarkedCount = 0;
            // Số lượng học sinh "vắng" (xóa x / để trống ô điểm danh).
            var absentClearedCount = 0;
            // Số dòng trên sheet không map được học sinh trong DB.
            var unmatchedSheetRowCount = 0;
            // Có sử dụng fallback full name hay không.
            var usedFallbackNameMatch = false;

            // Duyệt từng dòng học sinh trên sheet để map với học sinh trong attendance.
            foreach (var sheetRow in sheetRows)
            {
                // Thử map theo split name, nếu thất bại thì fallback full name.
                var student = TryMatchStudent(
                    sheetRow,
                    studentsBySplitName,
                    studentsByFullName,
                    consumedStudentIds,
                    out var matchedByFallback);

                // Nếu không map được thì bỏ qua dòng này.
                if (student == null)
                {
                    unmatchedSheetRowCount++;
                    continue;
                }

                // Đã map thành công 1 học sinh.
                matchedStudentCount++;
                // Ghi nhận có fallback hay không.
                usedFallbackNameMatch |= matchedByFallback;

                // Nếu vắng => để rỗng. Nếu có mặt => ghi "x".
                var attendanceMark = string.Equals(student.AttendanceStatus, AttendanceStatus.Absent, StringComparison.OrdinalIgnoreCase)
                    ? string.Empty
                    : PresenceMark;

                // Đếm thống kê có mặt.
                if (attendanceMark == PresenceMark)
                {
                    presentMarkedCount++;
                }
                // Đếm thống kê vắng.
                else
                {
                    absentClearedCount++;
                }

                // Chỉ cập nhật cột điểm danh (x/blank) cho dòng học sinh.
                // Không cập nhật cột xếp loại để tránh ghi đè dữ liệu đánh giá trên sheet.
                updates.Add(BuildSingleCellUpdate(
                    BuildRange(worksheetName, $"{targetColumnLetter}{sheetRow.RowNumber}"),
                    attendanceMark));
            }

            // Tạo request batch để gửi tất cả thay đổi trong 1 lần gọi API.
            var request = new BatchUpdateValuesRequest
            {
                // USER_ENTERED để Google Sheet xử lý giống người dùng nhập tay.
                ValueInputOption = "USER_ENTERED",
                // Toàn bộ ô cần cập nhật.
                Data = updates
            };

            // Gọi Google Sheets API để ghi dữ liệu lên sheet.
            await sheetsService.Spreadsheets.Values
                .BatchUpdate(request, spreadsheetId)
                .ExecuteAsync(cancellationToken);

            // Trả về metadata để FE/API biết đã ghi vào đâu và thống kê map.
            return new GoogleSheetSyncResult
            {
                SpreadsheetId = spreadsheetId,
                WorksheetName = worksheetName,
                ColumnIndex = targetColumnIndex,
                ColumnLetter = targetColumnLetter,
                MatchedStudentCount = matchedStudentCount,
                TotalStudentCount = attendance.Students.Count,
                PresentMarkedCount = presentMarkedCount,
                AbsentClearedCount = absentClearedCount,
                UnmatchedSheetRowCount = unmatchedSheetRowCount,
                UnmatchedAttendanceStudentCount = Math.Max(0, attendance.Students.Count - matchedStudentCount),
                UsedFallbackNameMatch = usedFallbackNameMatch
            };
        }

        private async Task<GoogleSheetSyncResult> SyncStudentMetadataToSheetAsync(
            SheetsService sheetsService,
            string spreadsheetId,
            string worksheetName,
            IReadOnlyCollection<StudentResponse> students,
            CancellationToken cancellationToken)
        {
            var (classificationColumnIndex, notesColumnIndex) = await ResolveStudentMetadataColumnsStableAsync(
                sheetsService,
                spreadsheetId,
                worksheetName,
                cancellationToken);

            var classificationColumnLetter = ToColumnLetter(classificationColumnIndex);
            var notesColumnLetter = ToColumnLetter(notesColumnIndex);

            var sheetRows = await LoadSheetStudentRowsAsync(
                sheetsService,
                spreadsheetId,
                worksheetName,
                cancellationToken);

            var studentsForLookup = students.Select(student => new ScheduleAttendanceStudentResponse
            {
                StudentId = student.Id,
                MiddleName = student.MiddleName ?? string.Empty,
                FirstName = student.FirstName ?? string.Empty,
                FullName = $"{student.MiddleName} {student.FirstName}".Trim(),
                CompetencyLevel = student.CompetencyLevel ?? string.Empty,
                Note = student.Notes ?? string.Empty
            }).ToList();

            var studentsBySplitName = BuildStudentLookup(studentsForLookup, useFullName: false);
            var studentsByFullName = BuildStudentLookup(studentsForLookup, useFullName: true);
            var consumedStudentIds = new HashSet<string>(StringComparer.Ordinal);

            var updates = new List<ValueRange>();
            var matchedStudentCount = 0;
            var unmatchedSheetRowCount = 0;
            var usedFallbackNameMatch = false;

            foreach (var sheetRow in sheetRows)
            {
                var student = TryMatchStudent(
                    sheetRow,
                    studentsBySplitName,
                    studentsByFullName,
                    consumedStudentIds,
                    out var matchedByFallback);

                if (student == null)
                {
                    unmatchedSheetRowCount++;
                    continue;
                }

                matchedStudentCount++;
                usedFallbackNameMatch |= matchedByFallback;

                var classification = string.IsNullOrWhiteSpace(student.CompetencyLevel)
                    ? string.Empty
                    : student.CompetencyLevel.Trim().ToUpperInvariant();
                var notes = student.Note?.Trim() ?? string.Empty;

                updates.Add(BuildSingleCellUpdate(
                    BuildRange(worksheetName, $"{classificationColumnLetter}{sheetRow.RowNumber}"),
                    classification));
                updates.Add(BuildSingleCellUpdate(
                    BuildRange(worksheetName, $"{notesColumnLetter}{sheetRow.RowNumber}"),
                    notes));
            }

            if (updates.Count > 0)
            {
                var request = new BatchUpdateValuesRequest
                {
                    ValueInputOption = "USER_ENTERED",
                    Data = updates
                };

                await sheetsService.Spreadsheets.Values
                    .BatchUpdate(request, spreadsheetId)
                    .ExecuteAsync(cancellationToken);
            }

            return new GoogleSheetSyncResult
            {
                SpreadsheetId = spreadsheetId,
                WorksheetName = worksheetName,
                ColumnIndex = classificationColumnIndex,
                ColumnLetter = classificationColumnLetter,
                ClassificationColumnIndex = classificationColumnIndex,
                ClassificationColumnLetter = classificationColumnLetter,
                NotesColumnIndex = notesColumnIndex,
                NotesColumnLetter = notesColumnLetter,
                MatchedStudentCount = matchedStudentCount,
                TotalStudentCount = students.Count,
                UnmatchedSheetRowCount = unmatchedSheetRowCount,
                UnmatchedAttendanceStudentCount = Math.Max(0, students.Count - matchedStudentCount),
                UnmatchedStudentCount = Math.Max(0, students.Count - matchedStudentCount),
                UsedFallbackNameMatch = usedFallbackNameMatch
            };
        }

        private async Task<(int ClassificationColumnIndex, int NotesColumnIndex)> ResolveStudentMetadataColumnsStableAsync(
            SheetsService sheetsService,
            string spreadsheetId,
            string worksheetName,
            CancellationToken cancellationToken)
        {
            var request = sheetsService.Spreadsheets.Values.Get(
                spreadsheetId,
                BuildRange(worksheetName, "A5:ZZ5"));

            var response = await request.ExecuteAsync(cancellationToken);
            var rows = response.Values ?? new List<IList<object>>();
            var headerRow = rows.Count > 0 ? rows[0] : new List<object>();

            // Rule nghiep vu: cot xep loai cua bang hoc sinh luon la cot H.
            // Khong map sang KQ THI de tranh ghi nham vao cot tong ket.
            const int fixedClassificationColumnIndex = 8; // H
            var classificationColumnIndex = fixedClassificationColumnIndex;

            // Cot GHI CHU co the thay doi vi giao vien chen them cot buoi hoc.
            // Vi vay tim GHI CHU theo vi tri sau cum cot diem danh (B1/B2/OT/GMT...).
            var attendanceOffsets = Enumerable.Range(0, headerRow.Count)
                .Where(offset => IsAttendanceHeaderCell(GetCellValue(headerRow, offset)))
                .ToList();

            var notesColumnIndex = -1;
            if (attendanceOffsets.Count > 0)
            {
                var lastAttendanceOffset = attendanceOffsets.Max();
                for (var offset = lastAttendanceOffset + 1; offset < headerRow.Count; offset++)
                {
                    var header = NormalizeText(GetCellValue(headerRow, offset));
                    if (header.Contains("ghi chu", StringComparison.Ordinal))
                    {
                        notesColumnIndex = offset + 1;
                        break;
                    }
                }
            }

            // Fallback: neu khong tim duoc theo cum diem danh thi quet toan bo dong header.
            if (notesColumnIndex <= 0)
            {
                for (var offset = 0; offset < headerRow.Count; offset++)
                {
                    var header = NormalizeText(GetCellValue(headerRow, offset));
                    if (header.Contains("ghi chu", StringComparison.Ordinal))
                    {
                        notesColumnIndex = offset + 1;
                        break;
                    }
                }
            }

            if (notesColumnIndex <= 0)
            {
                throw new InvalidOperationException("Khong tim thay cot GHI CHU tren dong header (hang 5).");
            }

            return (classificationColumnIndex, notesColumnIndex);
        }

        private async Task<(int ClassificationColumnIndex, int NotesColumnIndex)> ResolveStudentMetadataColumnsAsync(
            SheetsService sheetsService,
            string spreadsheetId,
            string worksheetName,
            CancellationToken cancellationToken)
        {
            var request = sheetsService.Spreadsheets.Values.Get(
                spreadsheetId,
                BuildRange(worksheetName, "A5:ZZ5"));

            var response = await request.ExecuteAsync(cancellationToken);
            var rows = response.Values ?? new List<IList<object>>();
            var headerRow = rows.Count > 0 ? rows[0] : new List<object>();

            var notesColumnIndex = -1;
            var kqThiColumnIndex = -1;
            var xepLoaiColumnIndex = -1;
            var xlColumnIndex = -1;

            for (var offset = 0; offset < headerRow.Count; offset++)
            {
                var header = NormalizeText(GetCellValue(headerRow, offset));
                if (string.IsNullOrWhiteSpace(header))
                {
                    continue;
                }

                var columnIndex = offset + 1;

                if (notesColumnIndex < 0 && header.Contains("ghi chu", StringComparison.Ordinal))
                {
                    notesColumnIndex = columnIndex;
                }

                if (kqThiColumnIndex < 0 && (header == "kq thi" || header.Contains("ket qua thi", StringComparison.Ordinal)))
                {
                    kqThiColumnIndex = columnIndex;
                }

                if (xepLoaiColumnIndex < 0 && header.Contains("xep loai", StringComparison.Ordinal))
                {
                    xepLoaiColumnIndex = columnIndex;
                }

                if (xlColumnIndex < 0 && header == "xl")
                {
                    xlColumnIndex = columnIndex;
                }
            }

            var classificationColumnIndex = kqThiColumnIndex > 0
                ? kqThiColumnIndex
                : (xepLoaiColumnIndex > 0
                    ? xepLoaiColumnIndex
                    : xlColumnIndex);

            if (classificationColumnIndex <= 0)
            {
                throw new InvalidOperationException(
                    "Không tìm thấy cột xếp loại trên sheet (header: 'KQ THI', 'Xếp loại' hoặc 'XL').");
            }

            if (notesColumnIndex <= 0)
            {
                throw new InvalidOperationException(
                    "Không tìm thấy cột ghi chú trên sheet (header chứa 'GHI CHÚ').");
            }

            return (classificationColumnIndex, notesColumnIndex);
        }

        private async Task<int> ResolveTargetColumnIndexAsync(
            SheetsService sheetsService,
            string spreadsheetId,
            string worksheetName,
            ScheduleAttendanceResponse attendance,
            CancellationToken cancellationToken)
        {
            // Đọc 3 hàng header của khu vực điểm danh:
            // - Hàng 5: tiêu đề buổi/tiết
            // - Hàng 6: số tiết
            // - Hàng 7: ngày dạy
            // Bắt đầu từ cột I (I5) đến ZZ7 để có đủ buffer.
            var request = sheetsService.Spreadsheets.Values.Get(
                spreadsheetId,
                BuildRange(worksheetName, "I5:ZZ7"));

            // Gọi Google Sheets API để lấy dữ liệu header hiện tại.
            var response = await request.ExecuteAsync(cancellationToken);
            // Đảm bảo không null để xử lý an toàn.
            var rows = response.Values ?? new List<IList<object>>();

            // Tách từng hàng header. Nếu sheet không đủ hàng thì fallback list rỗng.
            var row5 = rows.Count > 0 ? rows[0] : new List<object>();
            var row6 = rows.Count > 1 ? rows[1] : new List<object>();
            var row7 = rows.Count > 2 ? rows[2] : new List<object>();

            // Số cột cần duyệt = độ dài lớn nhất trong 3 hàng header.
            // Vì từng hàng có thể bị thiếu cell ở cuối.
            var maxLength = new[] { row5.Count, row6.Count, row7.Count }.Max();
            // Ngày của lịch học cần đồng bộ (chỉ lấy phần Date).
            var scheduleDate = attendance.Date.Date;

            // Không có dữ liệu header nào -> mặc định ghi vào cột I.
            if (maxLength <= 0)
            {
                return AttendanceStartColumnIndex;
            }

            // Thử nhận diện "cụm cột điểm danh" dựa trên header hàng 5 (B1, OT1, GMT1, SL..., B. SUNG...).
            // Mục tiêu: tránh ghi nhầm vào các bảng/khung khác trong cùng worksheet (ví dụ cột CR).
            var attendanceHeaderOffsets = Enumerable.Range(0, maxLength)
                .Where(offset => IsAttendanceHeaderCell(GetCellValue(row5, offset)))
                .ToList();

            if (attendanceHeaderOffsets.Count > 0)
            {
                // A) Ưu tiên cột trong cụm điểm danh đã có đúng ngày.
                foreach (var offset in attendanceHeaderOffsets)
                {
                    var existingDateRaw = GetCellValue(row7, offset);
                    if (CellMatchesDate(existingDateRaw, scheduleDate))
                    {
                        return AttendanceStartColumnIndex + offset;
                    }
                }

                // B) Nếu chưa có ngày, append ở bên phải theo cột có ngày gần nhất.
                // Không quay lại điền vào "lỗ hổng" ở giữa để tránh nhảy cột bất ngờ.
                var usedDateOffsets = attendanceHeaderOffsets
                    .Where(offset => !string.IsNullOrWhiteSpace(GetCellValue(row7, offset)))
                    .ToList();

                if (usedDateOffsets.Count > 0)
                {
                    var lastUsedDateOffset = usedDateOffsets.Max();

                    // Nếu trong mẫu đã có cột dự phòng ở bên phải, dùng cột dự phòng đầu tiên sau cột cuối cùng đã dùng.
                    var nextTemplateOffset = attendanceHeaderOffsets
                        .Where(offset => offset > lastUsedDateOffset
                            && string.IsNullOrWhiteSpace(GetCellValue(row7, offset)))
                        .OrderBy(offset => offset)
                        .FirstOrDefault(-1);

                    // Nếu tìm thấy cột template trống nằm bên phải cột đã dùng gần nhất,
                    // thì ưu tiên ghi vào đó để giữ đúng khung mẫu sẵn có.
                    if (nextTemplateOffset >= 0)
                    {
                        return AttendanceStartColumnIndex + nextTemplateOffset;
                    }

                    // Nếu không còn cột dự phòng trong template -> ghi tiếp ngay sau cột đã dùng cuối cùng.
                    return AttendanceStartColumnIndex + lastUsedDateOffset + 1;
                }

                // C) Chưa có ngày nào trong cụm điểm danh -> bắt đầu tại cột đầu tiên của cụm.
                return AttendanceStartColumnIndex + attendanceHeaderOffsets.Min();
            }

            // 1) Ưu tiên tìm cột đã có đúng ngày học để cập nhật lại đúng cột cũ.
            // Mục tiêu: một ngày học chỉ map vào 1 cột cố định, tránh tạo cột mới không cần thiết.
            for (var offset = 0; offset < maxLength; offset++)
            {
                // Giá trị ngày đang có tại hàng 7 của cột hiện tại.
                var existingDateRaw = GetCellValue(row7, offset);
                // CellMatchesDate hỗ trợ match cả text, date locale và serial date.
                if (CellMatchesDate(existingDateRaw, scheduleDate))
                {
                    // Trả về index cột thực tế trong sheet (I + offset).
                    return AttendanceStartColumnIndex + offset;
                }
            }

            // 2) Nếu chưa có cột theo ngày, lấy cột đầu tiên có hàng ngày (hàng 7) đang trống.
            // Không yêu cầu giáo viên phải nhập tay ngày/số tiết trước trên Google Sheet.
            for (var offset = 0; offset < maxLength; offset++)
            {
                var existingDateRaw = GetCellValue(row7, offset);
                if (string.IsNullOrWhiteSpace(existingDateRaw))
                {
                    // Chọn cột trống để hệ thống tự điền hàng 5/6/7 và ghi điểm danh.
                    return AttendanceStartColumnIndex + offset;
                }
            }

            // 3) Nếu tất cả cột đang có đều đã có ngày, tạo cột mới ngay sau cột cuối cùng.
            // Lưu ý: đây là "ghi sang cột tiếp theo", không phải insert chen giữa các cột.
            return AttendanceStartColumnIndex + maxLength;
        }

        private static bool IsAttendanceHeaderCell(string rawHeader)
        {
            if (string.IsNullOrWhiteSpace(rawHeader))
            {
                return false;
            }

            // Normalize để so khớp ổn định với chữ hoa/thường, dấu tiếng Việt, dấu câu.
            var normalized = NormalizeText(rawHeader);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            // Các pattern phổ biến trong mẫu cột điểm danh.
            if (Regex.IsMatch(normalized, @"^(b\s*\d+|ot\s*\d+|gmt\s*\d+)$"))
            {
                return true;
            }

            if (normalized.StartsWith("sl ", StringComparison.Ordinal)
                || normalized == "sl"
                || normalized.Contains("b sung", StringComparison.Ordinal)
                || normalized.Contains("bo sung", StringComparison.Ordinal))
            {
                return true;
            }

            return false;
        }

        private static bool CellMatchesDate(string rawValue, DateTime targetDate)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return false;
            }

            // Trường hợp sheet lưu ngày dạng text: dd/MM/yyyy, d/M/yyyy, ...
            var normalized = NormalizeText(rawValue);
            var acceptedTokens = BuildDateMatchTokens(targetDate);
            if (acceptedTokens.Contains(normalized))
            {
                return true;
            }

            // Trường hợp sheet lưu ngày dạng date/chuỗi date locale.
            if (DateTime.TryParse(rawValue, out var parsedDate))
            {
                return parsedDate.Date == targetDate.Date;
            }

            // Trường hợp sheet lưu serial date (số OADate).
            if (double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var serialValue)
                || double.TryParse(rawValue, NumberStyles.Float, CultureInfo.CurrentCulture, out serialValue))
            {
                try
                {
                    var serialDate = DateTime.FromOADate(serialValue).Date;
                    return serialDate == targetDate.Date;
                }
                catch
                {
                    // Serial không hợp lệ -> bỏ qua và xem là không match.
                }
            }

            return false;
        }

        private async Task<List<SheetStudentRow>> LoadSheetStudentRowsAsync(
            SheetsService sheetsService,
            string spreadsheetId,
            string worksheetName,
            CancellationToken cancellationToken)
        {
            var request = sheetsService.Spreadsheets.Values.Get(
                spreadsheetId,
                BuildRange(worksheetName, $"B{StudentStartRow}:C"));
            var response = await request.ExecuteAsync(cancellationToken);
            var values = response.Values ?? new List<IList<object>>();

            var rows = new List<SheetStudentRow>(values.Count);
            for (var index = 0; index < values.Count; index++)
            {
                var row = values[index];
                var middleName = GetCellValue(row, 0);
                var firstName = GetCellValue(row, 1);
                var splitNameKey = BuildStudentNameKey(middleName, firstName);
                var fullNameKey = BuildFullNameKey(middleName, firstName);
                if (string.IsNullOrWhiteSpace(splitNameKey) && string.IsNullOrWhiteSpace(fullNameKey))
                {
                    continue;
                }

                rows.Add(new SheetStudentRow(StudentStartRow + index, splitNameKey, fullNameKey));
            }

            return rows;
        }

        private static Dictionary<string, Queue<ScheduleAttendanceStudentResponse>> BuildStudentLookup(
            IReadOnlyCollection<ScheduleAttendanceStudentResponse> students,
            bool useFullName)
        {
            return students
                .GroupBy(s => useFullName
                    ? BuildFullNameKey(s.MiddleName, s.FirstName)
                    : BuildStudentNameKey(s.MiddleName, s.FirstName))
                .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                .ToDictionary(
                    g => g.Key,
                    g => new Queue<ScheduleAttendanceStudentResponse>(g.OrderBy(x => x.StudentId)));
        }

        private static ScheduleAttendanceStudentResponse? TryMatchStudent(
            SheetStudentRow sheetRow,
            Dictionary<string, Queue<ScheduleAttendanceStudentResponse>> bySplitName,
            Dictionary<string, Queue<ScheduleAttendanceStudentResponse>> byFullName,
            HashSet<string> consumedStudentIds,
            out bool matchedByFallback)
        {
            matchedByFallback = false;

            var matchedBySplit = TryDequeueStudentByKey(
                bySplitName,
                sheetRow.SplitNameKey,
                consumedStudentIds);
            if (matchedBySplit != null)
            {
                return matchedBySplit;
            }

            var matchedByFullName = TryDequeueStudentByKey(
                byFullName,
                sheetRow.FullNameKey,
                consumedStudentIds);
            if (matchedByFullName != null)
            {
                matchedByFallback = true;
                return matchedByFullName;
            }

            return null;
        }

        private static ScheduleAttendanceStudentResponse? TryDequeueStudentByKey(
            Dictionary<string, Queue<ScheduleAttendanceStudentResponse>> lookup,
            string key,
            HashSet<string> consumedStudentIds)
        {
            if (string.IsNullOrWhiteSpace(key)
                || !lookup.TryGetValue(key, out var queue)
                || queue.Count == 0)
            {
                return null;
            }

            while (queue.Count > 0)
            {
                var candidate = queue.Dequeue();
                if (string.IsNullOrWhiteSpace(candidate.StudentId))
                {
                    continue;
                }

                if (consumedStudentIds.Add(candidate.StudentId))
                {
                    return candidate;
                }
            }

            return null;
        }

        private static ValueRange BuildSingleCellUpdate(string range, object value)
        {
            return new ValueRange
            {
                Range = range,
                Values = new List<IList<object>>
                {
                    new List<object> { value }
                }
            };
        }

        private static int ResolvePeriodCount(ScheduleAttendanceResponse attendance)
        {
            if (TryResolveCountFromPeriodLabel(attendance.PeriodLabel, out var countByLabel))
            {
                return countByLabel;
            }

            if (TimeSpan.TryParse(attendance.StartTime, out var start)
                && TimeSpan.TryParse(attendance.EndTime, out var end)
                && end > start)
            {
                var lessonCount = (int)Math.Round((end - start).TotalMinutes / 45.0, MidpointRounding.AwayFromZero);
                return Math.Max(1, lessonCount);
            }

            return 1;
        }

        private static bool TryResolveCountFromPeriodLabel(string? periodLabel, out int count)
        {
            count = 1;
            if (string.IsNullOrWhiteSpace(periodLabel))
            {
                return false;
            }

            var normalized = periodLabel.Trim();
            var rangeMatch = Regex.Match(normalized, @"(?<start>\d+)\s*-\s*(?<end>\d+)");
            if (rangeMatch.Success)
            {
                var start = int.Parse(rangeMatch.Groups["start"].Value);
                var end = int.Parse(rangeMatch.Groups["end"].Value);
                count = Math.Abs(end - start) + 1;
                return true;
            }

            var numbers = Regex.Matches(normalized, @"\d+")
                .Select(m => m.Value)
                .ToList();

            if (numbers.Count > 1)
            {
                count = numbers.Count;
                return true;
            }

            if (numbers.Count == 1)
            {
                count = 1;
                return true;
            }

            return false;
        }

        private static HashSet<string> BuildDateMatchTokens(DateTime date)
        {
            var tokens = new[]
            {
                date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture),
                date.ToString("d/M/yyyy", CultureInfo.InvariantCulture),
                date.ToString("dd/MM", CultureInfo.InvariantCulture),
                date.ToString("d/M", CultureInfo.InvariantCulture),
            };

            return tokens
                .Select(NormalizeText)
                .ToHashSet(StringComparer.Ordinal);
        }

        private static string BuildStudentNameKey(string? middleName, string? firstName)
        {
            var middle = NormalizeText(middleName);
            var first = NormalizeText(firstName);
            if (string.IsNullOrWhiteSpace(middle) && string.IsNullOrWhiteSpace(first))
            {
                return string.Empty;
            }

            return $"{middle}|{first}";
        }

        private static string BuildFullNameKey(string? middleName, string? firstName)
        {
            return NormalizeText($"{middleName} {firstName}");
        }

        private static string NormalizeText(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var normalized = value.Trim().ToLowerInvariant().Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(normalized.Length);

            foreach (var c in normalized)
            {
                var category = CharUnicodeInfo.GetUnicodeCategory(c);
                if (category != UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(c);
                }
            }

            var noDiacritics = sb.ToString()
                .Normalize(NormalizationForm.FormC)
                .Replace('đ', 'd')
                .Replace('Đ', 'D');
            noDiacritics = noDiacritics
                .Replace('\u0111', 'd')
                .Replace('\u0110', 'D');
            var normalizedPunctuation = Regex.Replace(noDiacritics, @"[^\p{L}\p{N}\s]", " ");
            return Regex.Replace(normalizedPunctuation, @"\s+", " ").Trim();
        }

        private static string GetCellValue(IList<object> row, int index)
        {
            if (row.Count <= index || row[index] == null)
            {
                return string.Empty;
            }

            return row[index]?.ToString()?.Trim() ?? string.Empty;
        }

        private static string BuildRange(string worksheetName, string cellRange)
        {
            var escaped = worksheetName.Replace("'", "''");
            return $"'{escaped}'!{cellRange}";
        }

        private static string ToColumnLetter(int columnIndex)
        {
            if (columnIndex <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }

            var column = columnIndex;
            var letters = string.Empty;
            while (column > 0)
            {
                column--;
                letters = (char)('A' + (column % 26)) + letters;
                column /= 26;
            }

            return letters;
        }

        private static string? ParseSpreadsheetId(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            var raw = value.Trim();
            var urlMatch = Regex.Match(raw, @"spreadsheets\/d\/([a-zA-Z0-9\-_]+)", RegexOptions.IgnoreCase);
            if (urlMatch.Success)
            {
                return urlMatch.Groups[1].Value;
            }

            var idMatch = Regex.Match(raw, @"^[a-zA-Z0-9\-_]{20,}$");
            return idMatch.Success ? raw : null;
        }

        private static string FirstNonEmpty(params string?[] values)
        {
            foreach (var value in values)
            {
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }

            return string.Empty;
        }

        private SheetsService? BuildSheetsService(GoogleSheetsSettings settings)
        {
            GoogleCredential? credential = null;

            if (!string.IsNullOrWhiteSpace(settings.ServiceAccountJson))
            {
                var jsonBytes = Encoding.UTF8.GetBytes(settings.ServiceAccountJson.Trim());
                credential = GoogleCredential.FromStream(new MemoryStream(jsonBytes));
            }
            else if (!string.IsNullOrWhiteSpace(settings.ServiceAccountJsonPath))
            {
                var path = settings.ServiceAccountJsonPath.Trim();
                if (!Path.IsPathRooted(path))
                {
                    path = Path.Combine(AppContext.BaseDirectory, path);
                }

                if (!File.Exists(path))
                {
                    _logger.LogWarning("Google service account file not found: {Path}", path);
                    return null;
                }

                credential = GoogleCredential.FromFile(path);
            }

            if (credential == null)
            {
                return null;
            }

            if (credential.IsCreateScopedRequired)
            {
                credential = credential.CreateScoped(SheetsService.Scope.Spreadsheets);
            }

            return new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = string.IsNullOrWhiteSpace(settings.ApplicationName)
                    ? "MOS Grader"
                    : settings.ApplicationName.Trim()
            });
        }

        private sealed record SheetStudentRow(int RowNumber, string SplitNameKey, string FullNameKey);
    }
}
