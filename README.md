# Excel Converter v1.2.0

Chuyển hàng loạt **XLS / XLSM / XLSB → XLSX** bằng Microsoft Excel desktop trên Windows, hỗ trợ đường dẫn local dài trong thư mục OneDrive và SharePoint đã Sync.

## Chạy ứng dụng

Yêu cầu: Windows 10/11, Windows PowerShell 5.1 (hoặc PowerShell 7 để chạy CLI/test), Microsoft Excel desktop đã cài và kích hoạt. Giải nén toàn bộ gói ứng dụng vào thư mục local ngắn, ví dụ `C:\Tools\ExcelConverter`; giữ các file PS1, PSM1 và C# cạnh nhau. Không cần chạy Administrator.

1. Mở `ExcelConverter.cmd` hoặc `ExcelConverter.bat`.
2. **Add Folder**, **Files**, hoặc kéo thả nhiều file/thư mục từ Explorer vào queue.
3. Chọn **Include subfolders** nếu cần. Thư mục con `old` và file khóa `~$` được bỏ qua. Có thể chọn trực tiếp `old` nếu thực sự muốn xử lý thư mục đó.
4. Nhấn **CONVERT**. Giữ nguồn là mặc định; Move/Delete chỉ thực hiện sau khi xác minh kết quả. Timeout mặc định **180 giây/file**, có thể điều chỉnh.
5. **CANCEL** dừng file hiện tại và phần còn lại. **Reports** mở CSV kết quả tại `%LOCALAPPDATA%\ExcelConverter\Reports`.

File đã có đuôi `.xlsx` được ghi **AlreadyXlsx**, không tạo thêm bản chuyển đổi; đây là phân loại đuôi file, không phải chứng nhận file sạch/hợp lệ. Tên `abc.xlsm.xlsx` vẫn là XLSX. **Giải nén ZIP/7z trước** khi thêm dữ liệu vào queue.

## Đường dẫn OneDrive / SharePoint Sync

Chọn **đường dẫn local trong Explorer** của OneDrive hoặc thư viện SharePoint đã Sync. Ứng dụng dùng Unicode filesystem IO để đọc đường dẫn dài, copy workbook vào thư mục tạm ngắn, chuyển đổi bằng Excel rồi lưu kết quả cạnh nguồn. File trùng tên được tăng hậu tố, không ghi đè.

Placeholder cần sync client đang chạy, đủ quyền và có mạng để hydrate. File không tải được, bị khóa hoặc Excel treo được ghi lỗi/timeout riêng; batch tiếp tục. Nếu dữ liệu lớn, tăng timeout hoặc chọn **Always keep on this device** trước khi chạy.

Nếu thư mục Temp của profile quá dài/chứa dấu ngoặc vuông, đặt `EXCELCONVERTER_TEMP` tới thư mục local ngắn mà tài khoản có quyền ghi, ví dụ `C:\ECtemp`, rồi chạy lại. Không đặt thư mục tạm này trong nơi sync.

Giới hạn sync cloud vẫn áp dụng. Việc convert/lưu local không bảo đảm OneDrive đã upload thành công. Không hỗ trợ URL SharePoint trực tiếp, file chưa Sync, bypass quyền, mật khẩu hay DLP/IRM.

## Batch không giao diện

Chạy từ PowerShell trong thư mục ứng dụng:

```powershell
.\ConvertXLS_Batch.ps1 -Path 'C:\Users\User\OneDrive - Company\Reports', 'D:\Archive' -Recurse -TimeoutSeconds 300 -ReportPath 'C:\Temp\results.csv'
```

Mặc định `-OriginalAction Keep`. Có thể chọn `Move` (vào `old`) hoặc `Delete` (xóa vĩnh viễn sau xác minh). Không chỉnh sửa nguồn đồng thời với conversion; ưu tiên Keep khi thư mục có nhiều người cùng sửa. File có external workbook links luôn giữ nguồn và báo cảnh báo cần kiểm tra link tương đối.

Trạng thái CSV: `Converted`, `ConvertedWithWarning`, `AlreadyXlsx`, `ScanError`, `Failed`, `TimedOut`, `Canceled`, `NotProcessed`, `OutputSaved`. `OutputSaved` nghĩa là đã công bố output nhưng chưa xác nhận cleanup nguồn. CLI trả exit code 1 nếu có trạng thái cần kiểm tra.

## Kiểm tra và bảo toàn dữ liệu

- Mở nguồn read-only, không cập nhật external workbook links, tắt events và VBA khi automation mở file.
- Excel mở lại kết quả, kiểm tra định dạng; package XLSX/XML được kiểm tra trước công bố.
- Copy kết quả sang file tạm cạnh đích, so SHA-256, rename không ghi đè, sau đó mới xử lý bản gốc.
- Worker và Excel riêng theo file; timeout/hủy chỉ cleanup Excel PID/start time do worker ghi nhận sở hữu.
- **XLSX loại bỏ VBA**; macro UDF và tính năng legacy có thể không còn hoạt động. Kiểm tra kết quả trước khi dùng thay thế nguồn.
- Đây là công cụ conversion, **không phải antivirus**. AutomationSecurity không vô hiệu hóa XLM trong mọi trường hợp; Trust Center và chính sách doanh nghiệp vẫn áp dụng. Converter không tự xóa file theo tên nghi là virus và không đụng XLSTART.

Công cụ `XLSStartCleaner` cũ vẫn tồn tại độc lập trong repository, không đi kèm gói converter mới; không nằm trong luồng chuyển đổi mới. Heuristic tên file/module của nó có thể báo nhầm; xem mã và sao lưu trước khi sử dụng.

## Kiểm thử

```powershell
.\tests\Run.Tests.ps1
# Windows có Excel desktop: kiểm thử tích hợp với fixture tự tạo
.\tests\Excel.Smoke.Tests.ps1
```

CI cấu hình Windows PowerShell 5.1 và PowerShell 7. Kiểm thử trên macOS chỉ kiểm tra core/worker orchestration; không thay thế việc chạy Excel COM, WinForms và sync client thực trên Windows. [Báo cáo deep audit](docs/AUDIT-2026-10-26.md) ghi đầy đủ nguyên nhân, bằng chứng và giới hạn kiểm chứng.

---

## Tác giả

**Le Van An** (Vietnam IT)

[![GitHub](https://img.shields.io/badge/GitHub-anlvdt-181717?style=for-the-badge&logo=github)](https://github.com/anlvdt)

## Ủng hộ dự án

Nếu bạn thấy dự án hữu ích, hãy cân nhắc ủng hộ tác giả.

### Chuyển khoản

| Phương thức | Số tài khoản | Chủ tài khoản |
|------------|-------------|---------------|
| MB Bank | `0360126996868` | LE VAN AN |
| Momo | `0976896621` | LE VAN AN |

### Shopee Affiliate

Mình làm Affiliate Shopee, nếu thấy sản phẩm hữu ích hãy ủng hộ mình một click nhé. Chỉ cần click không cần mua cũng được!

[![Shopee](https://img.shields.io/badge/Shopee-EE4D2D?style=for-the-badge&logo=shopee&logoColor=white)](https://s.shopee.vn/7AYWh5NzOB)

**[Xem sản phẩm trên Shopee](https://s.shopee.vn/7AYWh5NzOB)** — Xin cảm ơn!

### Ủng hộ khác

- Star repo trên GitHub
- Chia sẻ dự án cho bạn bè, đồng nghiệp
- Báo bug hoặc đề xuất tính năng mới qua Issues
- Đóng góp code qua Pull Requests

---

## License

MIT License — Copyright (c) 2026 Le An (Vietnam IT)
