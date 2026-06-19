# Assignment -> Publication QA Checklist

Checklist này dùng cho phase siết rule nghiệp vụ `OTTH + Ôn Thi` sau refactor assignment-driven scheduling.

## Chuẩn bị

- Có ít nhất 1 lớp test với học sinh active.
- Có sẵn assignment thuộc 4 nhóm:
  - `OTTH` hợp lệ
  - `Ôn Thi` hợp lệ
  - assignment inactive
  - assignment legacy hoặc metadata lỗi
- Nếu cần rà dữ liệu trước, chạy:
 - Nếu cần rà dữ liệu trước, chạy:

```powershell
cd BACKEND
dotnet run --project .\tools\AssignmentPublicationAudit\AssignmentPublicationAudit.csproj -- --include-inactive --preview 50
```

- Nếu cần backfill metadata legacy sau khi rà, chạy dry-run trước:

```powershell
cd BACKEND
dotnet run --project .\tools\AssignmentMetadataBackfill\AssignmentMetadataBackfill.csproj -- --include-inactive --preview 50
```

## Case 1: Assignment OTTH hợp lệ tạo lịch thi được

- Mở màn tạo lịch thi trong `GradingView`.
- Chọn lớp, chọn học sinh, chọn 1 hoặc nhiều assignment `OTTH` có trạng thái dùng được.
- Tạo publication.
- Kỳ vọng:
  - backend trả thành công
  - publication có `assignmentIds[]`
  - `projectSequence[].sourceAssignmentId` được snapshot đúng
  - link thi/publication token hiển thị bình thường

## Case 2: Assignment Ôn Thi hợp lệ tạo lịch thi được

- Tạo assignment từ template `Ôn Thi`.
- Quay lại màn tạo lịch thi và chọn assignment vừa tạo.
- Tạo publication.
- Kỳ vọng:
  - publication tạo thành công
  - assignment `examType=onthi` vẫn chạy qua OTTH grader
  - UI không báo block reason

## Case 3: Assignment inactive bị chặn đúng

- Lấy một assignment đã inactive.
- Mở màn tạo lịch thi.
- Kỳ vọng:
  - assignment không chọn được
  - UI hiển thị lý do rõ ràng
  - nếu cố gọi API trực tiếp bằng `assignmentIds[]`, backend trả `400`

## Case 4: Assignment legacy / thiếu metadata bị chặn đúng

- Dùng assignment thiếu một trong các field:
  - `examType`
  - `subject`
  - `projectCode`
  - hoặc `gradingApiEndpoint` với assignment auto
- Kỳ vọng:
  - UI hiện trạng thái không publish được
  - `publishBlockReason` dễ hiểu
  - API tạo publication trả lỗi business rõ ràng, không fallback ngầm

## Case 5: Duplicate assignment bị chặn đúng

- Chọn cùng một assignment 2 lần trong payload `assignmentIds[]` bằng API hoặc tool test.
- Kỳ vọng:
  - backend trả `400`
  - message nêu rõ không được chọn trùng assignment trong cùng một lịch thi

## Case 6: Assignment khác lớp bị chặn đúng

- Cố tạo publication cho lớp A nhưng nhét assignment thuộc lớp B.
- Kỳ vọng:
  - backend trả `400`
  - message nêu rõ assignment phải thuộc đúng lớp

## Case 7: Assignment đã dùng publication vẫn sửa field an toàn được

- Chọn một assignment đã từng được dùng tạo publication.
- Sửa `name`, `description`, hoặc `isActive`.
- Kỳ vọng:
  - update thành công
  - publication cũ không bị thay đổi snapshot

## Case 8: Assignment đã dùng publication không sửa được field lõi

- Với cùng assignment trên, thử sửa:
  - `examType`
  - `subject`
  - `projectCode`
  - `gradingType`
  - `gradingApiEndpoint`
- Kỳ vọng:
  - backend chặn
  - UI báo lỗi rõ ràng

## Case 9: Publication cũ vẫn chạy bình thường

- Mở một publication đã tạo trước đó.
- Thực hiện bootstrap/public flow hoặc kiểm tra session hiện có.
- Kỳ vọng:
  - publication cũ vẫn đọc được
  - session/runtime không bị ảnh hưởng bởi việc assignment gốc đã đổi tên hoặc inactive

## Ghi nhận lỗi nếu fail

- `classId`
- `assignmentId`
- `examType`
- `subject`
- `projectCode`
- `gradingApiEndpoint`
- bước tái hiện
- message trả về từ backend
- ảnh chụp UI nếu reason hiển thị sai hoặc thiếu
