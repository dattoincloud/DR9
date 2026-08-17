# Checklist kiểm thử

## 1. Môi trường

- [ ] `scripts/smoke_test.vbs` trả PASS bằng đúng bitness dự kiến.
- [ ] `GET /` trả HTML với `charset=utf-8`.
- [ ] Handler `ASPClassic` xử lý `.asp`; source VBScript không bị trả thẳng về browser.
- [ ] `GET /App_Data/tasks.mdb` bị IIS từ chối.
- [ ] `GET /includes/db.asp` và `GET /scripts/create_database.vbs` bị IIS từ chối.
- [ ] Khi ghi dữ liệu, `.ldb` được tạo trong `App_Data` và biến mất sau khi connection đóng.
- [ ] App pool có một worker process và bitness khớp provider.

## 2. CRUD và HTTP

- [ ] Danh sách hiển thị 3 seed rows.
- [ ] Tìm theo title, description và assignee.
- [ ] Lọc đủ `Todo`, `InProgress`, `Done`; status giả trả HTTP 400.
- [ ] Tạo công việc rồi refresh trang danh sách không tạo bản ghi thứ hai.
- [ ] Double-click nút lưu nhanh chỉ gửi một form submit trong browser có JavaScript.
- [ ] Sửa title/status/due date; dữ liệu được load lại đúng.
- [ ] Xóa có JavaScript confirm; tắt JavaScript vẫn xóa được bằng POST.
- [ ] `GET /task_save.asp` và `GET /task_delete.asp` trả HTTP 405.
- [ ] ID rỗng, chữ, số âm, số quá dài và ID không tồn tại trả 400/404 phù hợp.
- [ ] CSRF token rỗng/sai trả HTTP 403.

## 3. Validation và bảo mật

- [ ] Title rỗng hoặc chỉ có khoảng trắng bị chặn server-side.
- [ ] Title chỉ có tab/newline hoặc khoảng trắng full-width `　` bị chặn.
- [ ] Kiểm tra biên 150/151 ký tự title, 100/101 assignee, 4000/4001 description.
- [ ] Priority `0`, `4`, `1.5`, `abc` bị chặn.
- [ ] Status ngoài allowlist bị chặn.
- [ ] Ngày `2028-02-29` hợp lệ; `2027-02-29`, `2026-02-30`, `16/08/2026` bị chặn.
- [ ] Search/title `x' OR 1=1--` không thay đổi cấu trúc query.
- [ ] Search literal `[`, `%`, `_` không gây lỗi và không bị hiểu thành wildcard.
- [ ] Title `<img src=x onerror=alert(1)>` hiển thị thành text, không execute.
- [ ] Title/assignee chứa `& < > " '` không phá HTML/attribute.

## 4. Encoding

Tạo, sửa, search rồi xóa từng dữ liệu sau:

- [ ] `備品ノートPC`
- [ ] `Máy tính xách tay`
- [ ] `東京・Hà Nội・O'Brien`
- [ ] Kana half-width/full-width.
- [ ] Ký tự supplementary như `𠮷` nếu nghiệp vụ yêu cầu; ghi rõ `Len` VBScript có thể đếm hai UTF-16 code units.

Sau validation fail, dữ liệu render lại form vẫn phải đúng Unicode.

## 5. Concurrency và recovery

- [ ] Hai browser tạo hai task khác nhau cùng lúc.
- [ ] Hai browser sửa cùng một task để ghi nhận hành vi “last write wins”.
- [ ] Tạo bản backup `.mdb`, thử restore khi site đã dừng.
- [ ] Không copy/replace `.mdb` lúc IIS còn connection mở.
- [ ] Ghi lại OS build, IIS features, provider/version, x86/x64, app-pool identity và ACL đã dùng.
