# Legacy Task Board — Classic ASP + MS Access

Demo CRUD nhỏ để học và mô phỏng đúng stack legacy thường gặp: HTML/CSS/JavaScript thuần, Classic ASP viết bằng VBScript, ADO và Microsoft Access `.mdb`.

Ứng dụng có giao diện Việt–Nhật, danh sách/thống kê, tìm kiếm, lọc trạng thái, tạo, sửa và xóa công việc. Database mẫu đã nằm tại `App_Data/tasks.mdb` với 3 bản ghi seed.

> Đây là môi trường học và bảo trì legacy. Microsoft nói Access không được thiết kế cho tải lớn, còn Access Database Engine mới cũng không được khuyến nghị làm backend cho ứng dụng web server-side. Không dùng kiến trúc này để khởi tạo production mới.

## Chạy nhanh trên Windows

Classic ASP không chạy bằng cách double-click file hoặc dùng VS Code Live Server. Nó cần IIS và feature `ASP`.

Mở **Windows PowerShell as Administrator**, chuyển vào thư mục project rồi chạy:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\setup-iis.ps1
```

Script sẽ:

- bật IIS, Classic ASP, ISAPI Extensions và Request Filtering;
- tạo site `DR9ClassicAspDemo` tại `http://127.0.0.1:8088/`;
- tạo app pool `No Managed Code`, một worker, chạy 32-bit để dùng Jet;
- cho Anonymous Authentication chạy bằng Application Pool Identity;
- chỉ cấp `Modify` cho app pool trên `App_Data`;
- giữ database và include/script khỏi truy cập trực tiếp qua HTTP.

Sau khi script hoàn tất, mở:

```text
http://127.0.0.1:8088/
```

Nếu Windows yêu cầu restart, restart máy rồi chạy lại cùng lệnh. Script có thể chạy lại an toàn với cùng site/path.

### Provider/bitness cố định

Demo cố ý dùng pool 32-bit + `.mdb` + `Microsoft.Jet.OLEDB.4.0` để môi trường tái lập đúng stack legacy, thay vì âm thầm fallback sang engine khác:

```powershell
.\scripts\setup-iis.ps1
```

Nếu hệ thống thật dùng `.accdb`/ACE, hãy tạo một cấu hình riêng với `Microsoft.ACE.OLEDB.12.0`, đổi file database và đặt app-pool bitness khớp đúng bản ACE đã cài. Không nên fallback qua nhiều provider cho mọi lỗi vì lỗi permission/corruption có thể bị che bởi lỗi provider kế tiếp.

## Kiểm tra database không cần IIS

Database trong repo được tạo và smoke-test thành công bằng Jet 32-bit. Có thể chạy lại bài test không để lại dữ liệu:

```powershell
& "$env:WINDIR\SysWOW64\cscript.exe" //nologo .\scripts\smoke_test.vbs
```

Kết quả mong đợi:

```text
PASS: schema, CRUD transaction, parameterized queries and Unicode round-trip
Provider: Microsoft.Jet.OLEDB.4.0
```

Kiểm tra compile VBScript của toàn bộ page mà không cần IIS:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\check_asp_syntax.ps1
```

Kiểm thử luôn quá trình tạo một database sạch trong thư mục tạm rồi tự dọn:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\test_database_build.ps1
```

Để tạo database mới, trước tiên dừng site và đổi tên file cũ để còn khả năng phục hồi:

```powershell
Stop-Website DR9ClassicAspDemo
Rename-Item .\App_Data\tasks.mdb tasks.backup.mdb
& "$env:WINDIR\SysWOW64\cscript.exe" //nologo .\scripts\create_database.vbs
Start-Website DR9ClassicAspDemo
```

Script tạo DB cố ý từ chối ghi đè file hiện hữu.

## Cấu trúc và luồng xử lý

```text
Browser
  -> IIS / Classic ASP
  -> VBScript validation + ADODB.Command
  -> Jet 4.0 OLE DB provider
  -> App_Data/tasks.mdb
```

| File | Vai trò |
|---|---|
| `default.asp` | `GET`: danh sách, thống kê, search và filter |
| `task_form.asp` | `GET`: form tạo/sửa và load bản ghi bằng ID |
| `task_save.asp` | `POST`: validate rồi `INSERT`/`UPDATE`, sau đó redirect |
| `task_delete.asp` | `POST`: kiểm tra token rồi `DELETE` |
| `includes/util.asp` | encode HTML, parse ID/ngày, allowlist, CSRF, lỗi HTTP |
| `includes/db.asp` | mở ADO connection và tạo positional parameters |
| `App_Data/tasks.mdb` | Access 2000-format `.mdb`, bảng `Tasks` |
| `scripts/create_database.vbs` | ADOX/ADO tạo schema và seed data từ CLI |
| `scripts/smoke_test.vbs` | transaction test cho schema, Unicode và parameterized query |
| `scripts/check_asp_syntax.ps1` | ghép include/ASP blocks rồi dùng VBScript engine compile từng page |
| `scripts/test_database_build.ps1` | tạo DB sạch trong temp, chạy smoke test rồi dọn các file test cụ thể |
| `scripts/setup-iis.ps1` | cài/config IIS, app pool, binding và ACL |
| `web.config` | default document, security headers, chặn DB/include/script qua HTTP |

Request ghi dữ liệu theo mẫu Post/Redirect/Get, nên refresh trang danh sách không submit lặp.

## Những pattern legacy đáng học trong demo

### 1. ADO dùng dấu `?` theo vị trí

Access OLE DB không bind theo tên parameter trong SQL. Thứ tự `Parameters.Append` phải khớp tuyệt đối với thứ tự dấu `?`:

```asp
Set command = NewCommand(connection, _
    "UPDATE [Tasks] SET [Title] = ?, [Status] = ? WHERE [Id] = ?")
Call AddTextParameter(command, taskTitle, 150)
Call AddTextParameter(command, taskStatus, 20)
Call AddIntegerParameter(command, taskId)
```

Không nối `Request.Form` hoặc `Request.QueryString` vào SQL. Search trong `default.asp` cũng dùng parameters.

### 2. Validation phía server mới là nguồn quyết định

HTML `required`/`maxlength` và JavaScript chỉ cải thiện UX. `task_save.asp` vẫn kiểm tra bắt buộc, độ dài, allowlist status/priority và tự parse ngày ISO bằng `DateSerial` để không phụ thuộc locale máy.

### 3. Output phải encode

Mọi dữ liệu từ request/database đều đi qua `Server.HTMLEncode` (`H(...)`) trước khi render. Demo cũng dùng POST-only delete và token theo session để minh họa CSRF defense cơ bản.

### 4. Unicode phải thống nhất cả đường đi

Các page dùng `CodePage=65001`, response UTF-8 và ADO `adVarWChar`/`adLongVarWChar`. Smoke test ghi rồi đọc lại chuỗi Nhật–Việt trong transaction. Không trộn source UTF-8 với page code 932/Shift_JIS nếu chưa xác định rõ yêu cầu hệ thống thật.

### 5. Access cần quyền ghi cả thư mục

Jet tạo `.ldb` bên cạnh `.mdb`; chỉ cấp quyền ghi trên file database là chưa đủ. App pool cần `Modify` trên `App_Data`, nhưng chỉ cần `Read & Execute` trên phần code còn lại.

## Tự cấu hình IIS bằng giao diện

Nếu không muốn chạy script:

1. Mở **Turn Windows features on or off**.
2. Bật `Internet Information Services` → `World Wide Web Services` → `Application Development Features` → `ASP` và `ISAPI Extensions`.
3. Bật `Common HTTP Features` → `Default Document`, `Static Content`, `HTTP Errors`; bật `Request Filtering` và `IIS Management Console`.
4. Trong IIS Manager, tạo app pool riêng: `.NET CLR = No Managed Code`, pipeline `Integrated`, `Enable 32-Bit Applications = True`, `Load User Profile = True`, `Maximum Worker Processes = 1`.
5. Tạo website trỏ physical path vào root repo, binding local port `8088`.
6. Anonymous Authentication → **Edit** → **Application pool identity**.
7. Cấp `Read & Execute` cho `IIS AppPool\<PoolName>` ở root; chỉ cấp `Modify` tại `App_Data`.
8. Kiểm tra `http://127.0.0.1:8088/App_Data/tasks.mdb` bị từ chối (`404.x`), rồi kiểm thử CRUD.

Không cần bật ASP Parent Paths vì tất cả include đều dùng đường dẫn con tại root.

## Lỗi hay gặp

| Triệu chứng | Nguyên nhân thường gặp | Cách kiểm tra |
|---|---|---|
| `.asp` trả 404 hoặc không execute | Chưa bật ASP/ISAPI | Windows Features và IIS Handler Mappings `ASPClassic` |
| `Provider cannot be found` / `Class not registered` | Thiếu provider hoặc lệch x86/x64 | Jet → pool 32-bit; ACE phải khớp bitness |
| `Operation must use an updateable query` | Không tạo được `.ldb` | Cấp `Modify` cho đúng anonymous/app-pool identity trên `App_Data` |
| `Unspecified error 80004005` | ACL folder DB/temp hoặc profile | Xem IIS log; dùng Process Monitor lọc `w3wp.exe` + `ACCESS DENIED` |
| Chữ Nhật/Việt bị lỗi | Code page/source/provider type không thống nhất | UTF-8 + 65001 + `adVarWChar` |
| Chỉ thấy lỗi HTTP 500 chung | IIS ẩn chi tiết lỗi mặc định | Chỉ bật detailed ASP errors trên máy local, tắt lại sau debug |

## Giới hạn chủ động

- Không có login/phân quyền; đây là CRUD demo local.
- CSRF token dùng session và đủ để minh họa pattern, không phải thư viện security hiện đại.
- Access khóa theo file/page và không phù hợp nhiều request ghi đồng thời.
- Cập nhật đồng thời đang là “last write wins”; hệ thống thật nên có optimistic locking/version column.
- Không giữ `Connection`/`Recordset` trong `Session` hoặc `Application`; mỗi request mở, dùng rồi đóng.
- Trang list chủ ý chưa phân trang vì chỉ là demo; thêm pagination trước khi dùng với tập dữ liệu lớn.
- VBScript đã bị deprecated trên Windows; môi trường bảo trì cần ghi lại OS/features và giữ VM có thể tái tạo.
- JavaScript giữ cú pháp ES5-friendly, nhưng CSS của demo nhắm Edge/Chrome hiện đại. Nếu hệ thống thật bắt buộc IE11/IE mode, cần thêm compatibility stylesheet và test riêng trên đúng browser đó.

Xem checklist chi tiết tại [docs/TESTING.md](docs/TESTING.md).

## Tài liệu Microsoft

- [Build a Classic ASP website on IIS](https://learn.microsoft.com/en-us/iis/application-frameworks/running-classic-asp-applications-on-iis-7-and-iis-8/scenario-build-a-classic-asp-website-on-iis)
- [Using Classic ASP with Microsoft Access databases on IIS](https://learn.microsoft.com/en-us/iis/application-frameworks/running-classic-asp-applications-on-iis-7-and-iis-8/using-classic-asp-with-microsoft-access-databases-on-iis)
- [Microsoft 365 Access Runtime và giới hạn server-side](https://support.microsoft.com/en-US/Access/download-and-install-microsoft-365-access-runtime)
- [Access Database Engine 2016 lifecycle — đã hết extended support từ 10/2025](https://learn.microsoft.com/en-us/lifecycle/products/access-database-engine-2016-redistributable)
- [Windows deprecated features — VBScript](https://learn.microsoft.com/en-us/windows/whats-new/deprecated-features-resources#vbscript)
