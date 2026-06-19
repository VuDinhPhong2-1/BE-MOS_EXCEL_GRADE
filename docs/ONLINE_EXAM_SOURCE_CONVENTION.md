# Quy chuẩn bộ file nguồn cho Online Exam

## Mục tiêu
Cho mỗi project online chỉ cần đúng 3 file nguồn:

1. `template`
2. `instructions`
3. `help`

Không bắt buộc có `answer` nếu luồng chấm điểm không dựa vào so sánh file đáp án.

## Cấu trúc thư mục khuyến nghị

```text
OnlineExamSource/
  OTTH/
    Excel/
      P01/
        OTTH_EXCEL_P01_TEMPLATE_v1.xlsx
        OTTH_EXCEL_P01_INSTRUCTIONS_v1.txt
        OTTH_EXCEL_P01_HELP_v1.txt
    Word/
      P11/
        OTTH_WORD_P11_TEMPLATE_v1.docx
        OTTH_WORD_P11_INSTRUCTIONS_v1.txt
        OTTH_WORD_P11_HELP_v1.txt
  GMetrix/
    Excel/
    Word/
    PowerPoint/
```

## Quy tắc đặt tên file

Mẫu chung:

```text
{EXAMTYPE}_{SUBJECT}_{PROJECTCODE}_{KIND}_v{VERSION}.{EXT}
```

Ví dụ:

- `OTTH_EXCEL_P01_TEMPLATE_v1.xlsx`
- `OTTH_EXCEL_P01_INSTRUCTIONS_v1.txt`
- `OTTH_EXCEL_P01_HELP_v1.txt`
- `OTTH_WORD_P11_TEMPLATE_v1.docx`
- `OTTH_WORD_P11_INSTRUCTIONS_v1.txt`
- `OTTH_WORD_P11_HELP_v1.txt`

## Ý nghĩa từng file

### `TEMPLATE`
- Là file gốc học sinh sẽ mở và làm bài.
- Excel dùng `.xlsx` hoặc `.xlsm` nếu project cần macro.
- Word dùng `.docx`.
- PowerPoint dùng `.pptx`.

### `INSTRUCTIONS`
- Là toàn bộ đề bài cho project, lưu dạng text thuần `.txt`.
- Nội dung nên là phần mô tả chính thức mà Local Agent/online exam sẽ hiển thị cho học sinh.
- Nên viết theo block rõ ràng, mỗi yêu cầu một dòng hoặc một đoạn ngắn.

Ví dụ:

```text
Project 01 - Excel

1. Tính tổng cột Total cho toàn bộ bảng dữ liệu.
2. Lọc danh sách chỉ còn các agent có doanh số từ 30 trở lên.
3. Tạo biểu đồ cột cho vùng dữ liệu A4:H13.
```

### `HELP`
- Là nội dung hỗ trợ khi học sinh bấm Help, cũng dùng `.txt`.
- Không nên chép lại toàn bộ đề bài; nên là gợi ý thao tác, nhắc ribbon, hàm, menu, hoặc lưu ý lỗi hay gặp.

Ví dụ:

```text
Gợi ý:
- Dùng SUM để tính tổng.
- Dùng Filter trong tab Data để lọc dữ liệu.
- Nếu tạo chart, nhớ chọn đúng vùng dữ liệu trước khi vào Insert.
```

## Metadata trong hệ thống

- `projectCode` chỉ giữ identity học vụ của project, ví dụ `EXCEL_P01`, `WORD_P11`.
- Phân biệt OTTH, Ôn Thi, GMetrix bằng `examType`, không nhồi vào `projectCode`.
- `Ôn Thi` vẫn dùng source/grader OTTH, nên có thể tái sử dụng cùng bộ 3 file OTTH nếu cùng project.

## Mapping backend hiện tại

Backend hiện hỗ trợ upload assignment file theo `kind`:

- `template`
- `instructions`
- `help`
- `answer` vẫn giữ để tương thích dữ liệu cũ

Khi tạo `exam publication` từ assignment:

- `template` được snapshot thành `templateFileName`
- `instructions` được snapshot thành `instructionsFileName` + `instructionsText`
- `help` được snapshot thành `helpFileName` + `helpText`

Nhờ vậy publication/runtime không bị trôi nếu assignment bị sửa file sau này.
