# Chuyển PDF sang Word cho thầy Bảo

## Cài đặt dễ dàng

1. Máy tính cần có Microsoft Word bản desktop đã được kích hoạt.
2. Giữ nguyên thư mục ứng dụng đã nhận được.
3. Nháy đúp vào `CaiDatUngDungTiengViet.cmd` và làm theo thông báo.
4. Mở Start Menu và chọn **Chuyển PDF sang Word**.

Nếu Windows hỏi xác nhận, chỉ tiếp tục khi bạn nhận tệp từ nguồn đáng tin cậy. Việc cài đặt không cần quyền quản trị và chỉ tạo một lối tắt cho người dùng hiện tại.

## Cách dùng

1. Nhấn **Chọn tệp…** và chọn PDF hoặc tài liệu Word.
2. Kiểm tra thư mục lưu kết quả, rồi nhấn nút chuyển đổi tương ứng.
3. Khi chuyển PDF sang Word, ứng dụng tạo ba loại tệp:

   - `*.editable.docx`: tệp Word có thể chỉnh sửa.
   - `*.roundtrip.pdf`: PDF được xuất lại từ Word.
   - `*.fidelity-report.json`: báo cáo kỹ thuật về bố cục, phông chữ, bảng, hình, công thức và mức độ khác biệt.

4. Khi chuyển Word (`.docx`, `.docm` hoặc `.doc`) sang PDF, ứng dụng tạo:

   - `*.converted.pdf`: PDF chất lượng in do Microsoft Word xuất trực tiếp.
   - `*.word-to-pdf-report.json`: báo cáo về phông chữ, bảng, hình, công thức và văn bản. Với `.docx` hoặc `.docm`, báo cáo kiểm tra được đầy đủ hơn.

Nếu ứng dụng báo **cần kiểm tra**, hãy mở báo cáo (và ảnh so sánh khi có) trong thư mục kết quả. Luôn kiểm tra công thức, sơ đồ, bảng biểu, phần đầu/cuối trang và các ký tự tiếng Việt trước khi gửi tài liệu cho học sinh.

Tệp PDF gốc không bị thay đổi và ứng dụng không tải tài liệu lên Internet.
