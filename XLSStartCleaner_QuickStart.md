# 🛡️ XLSStart Cleaner - Quick Start Guide
# Hướng dẫn nhanh XLSStart Cleaner

---

## ⚡ Quick Start | Bắt đầu nhanh

### Cách 1: Từ Excel Converter (Khuyến nghị)
1. Chạy `ExcelConverter.cmd`
2. Nhấn nút **🛡 Scan XLSTART** ở góc trên
3. Chọn **SCAN ONLY** (chỉ quét) hoặc **SCAN & CLEAN** (quét và xóa)

### Cách 2: Chạy độc lập
- Double-click file `XLSStartCleaner.bat`

---

## 🔍 Công cụ này làm gì?

**XLSStart Cleaner** quét và xóa virus **X97M/Laroux** khỏi thư mục XLSTART của Excel.

### Virus X97M/Laroux là gì?
- Virus macro Excel cổ điển nhất
- Tự động lây nhiễm mỗi khi mở Excel
- Ẩn trong thư mục XLSTART
- Lan truyền sang tất cả file Excel bạn mở

### Công cụ này phát hiện:
- ✅ 27+ tên file virus đã biết
- ✅ 23+ tên module VBA độc hại
- ✅ 8+ hàm macro virus
- ✅ Quét 13 vị trí XLSTART (Office 2007-365)

---

## 🎯 Khi nào cần dùng?

### Dấu hiệu máy tính bị nhiễm virus Excel:
- ❌ Excel mở chậm bất thường
- ❌ Xuất hiện file lạ trong XLSTART
- ❌ Macro tự động chạy khi mở file
- ❌ File Excel tự động có macro dù không tạo
- ❌ Xuất hiện file: PERSONAL.XLS, VERA.XLS, STARTUP.XLS...

### Khuyến nghị:
- 🔄 Quét **1 lần/tuần** để phòng ngừa
- 🔄 Quét **ngay lập tức** nếu nghi ngờ nhiễm virus
- 🔄 Quét **sau khi mở file Excel từ nguồn không rõ**

---

## 📋 Hướng dẫn sử dụng

### SCAN ONLY (Quét thôi)
1. Nhấn nút **SCAN ONLY**
2. Đợi quét xong
3. Xem kết quả trong log
4. **Không xóa gì cả** - chỉ báo cáo

**Dùng khi:** Bạn muốn kiểm tra trước khi xóa

### SCAN & CLEAN (Quét và Xóa)
1. Nhấn nút **SCAN & CLEAN**
2. Xác nhận **YES** khi được hỏi
3. Đợi quét và xóa xong
4. **Xóa vĩnh viễn** file nhiễm virus

**Dùng khi:** Bạn chắc chắn muốn xóa virus

⚠️ **CẢNH BÁO:** Hành động này không thể hoàn tác!

---

## 📊 Hiểu kết quả quét

### Thông báo trong Log:

**[VIRUS]** - Phát hiện file nhiễm virus
```
[10:30:15] [VIRUS] PERSONAL.XLS - Known virus filename
```
→ File này là virus, cần xóa!

**[CLEAN]** - File an toàn
```
[10:30:16] [CLEAN] MyWorkbook.xls
```
→ File này OK, không có vấn đề

**[REMOVED]** - Đã xóa file virus
```
[10:30:17] [REMOVED] C:\...\XLSTART\VERA.XLS
```
→ File virus đã được xóa thành công

**[ERROR]** - Không thể xóa
```
[10:30:18] [ERROR] Could not remove: STARTUP.XLS - File may be in use
```
→ Đóng Excel và thử lại

---

## ❓ Xử lý sự cố

### Vấn đề: "Could not remove file - File may be in use"
**Giải pháp:**
1. Đóng TẤT CẢ cửa sổ Excel
2. Đóng các chương trình đang dùng Excel
3. Chạy lại công cụ

### Vấn đề: Không tìm thấy virus nhưng Excel vẫn lạ
**Giải pháp:**
1. Virus có thể nằm trong file workbook, không phải XLSTART
2. Dùng **Excel Converter** để chuyển đổi và làm sạch file
3. Quét virus bằng Windows Defender

### Vấn đề: Công cụ không chạy
**Giải pháp:**
1. Nhấp chuột phải `XLSStartCleaner.ps1` → **Run with PowerShell**
2. Kiểm tra Windows Defender có chặn không
3. Chạy từ Excel Converter chính

---

## 🔒 An toàn & Riêng tư

### Công cụ này:
- ✅ Chạy hoàn toàn OFFLINE
- ✅ KHÔNG gửi dữ liệu lên internet
- ✅ KHÔNG cài đặt gì vào máy
- ✅ KHÔNG sửa file Excel của bạn
- ✅ CHỈ quét thư mục XLSTART

### Công cụ này KHÔNG:
- ❌ Không truy cập file ngoài XLSTART
- ❌ Không gửi thông tin cá nhân
- ❌ Không cài virus/malware
- ❌ Không làm chậm máy tính

---

## 📚 Tài liệu đầy đủ

Xem file `XLSStartCleaner_UserGuide.md` để có hướng dẫn chi tiết.

---

## 💝 Ủng hộ tác giả

Nếu công cụ này hữu ích, hãy ủng hộ tác giả:

| Ngân hàng | Số tài khoản | Tên |
|-----------|--------------|-----|
| MB Bank | `0360126996868` | LE VAN AN |
| Momo | `0976896621` | LE VAN AN |

---

## 📞 Hỗ trợ

- 📧 Email: anlvdt@github
- 🌐 GitHub: github.com/anlvdt

---

**Giữ an toàn! Stay safe!** 🛡️

Version 1.0.0 | January 2026 | Le An (Vietnam IT)
