# 🖨️ Printer Manager (Portable)

![Giao diện Printer Manager](app.png)

**Công cụ Quản lý & Sửa lỗi Máy in Toàn diện trên Windows**  
*Phiên bản: 2.2.1 (Latest)*  
*Tác giả: @danhcp*

![Platform](https://img.shields.io/badge/Platform-Windows-0078D6.svg?style=flat-square)
![Type](https://img.shields.io/badge/Type-Portable_(EXE)-orange.svg?style=flat-square)
![License](https://img.shields.io/badge/License-Freeware-green.svg?style=flat-square)

## 📖 Giới thiệu

**Printer Manager** là phần mềm dạng **Portable** (chạy ngay không cần cài đặt), được thiết kế để thay thế giao diện quản lý máy in cũ kỹ của Windows bằng một giao diện tập trung, trực quan và mạnh mẽ hơn.

Công cụ giúp IT Helpdesk và người dùng văn phòng giải quyết nhanh các vấn đề:
*   Máy in bị kẹt lệnh, không thể xóa ("Remove device" không tác dụng).
*   Hệ thống chứa quá nhiều Driver rác gây xung đột.
*   Cần xem chi tiết ai đang in, file gì (Soi lệnh in).
*   Quản lý Share LAN và IP máy in tức thì.

## ✨ Tính năng nổi bật (v2.2.1)

### 🛠️ Sửa lỗi & Dọn dẹp Hệ thống
1.  **Xóa Máy In Tận Gốc:** Can thiệp Registry để xóa bỏ máy in cứng đầu (kèm tính năng **Auto Backup** Registry trước khi xóa để đảm bảo an toàn).
2.  **🧹 Dọn dẹp Driver Rác (MỚI):** Tự động quét và phát hiện các Driver (V3, V4) không còn được sử dụng bởi bất kỳ máy in nào và cho phép xóa chúng để giải phóng hệ thống.
3.  **Xóa Lệnh Kẹt (Clear Queue):** Xóa sạch toàn bộ lệnh in đang treo (Spool files) chỉ với 1 click.
4.  **Tự Động Xử Lý Spooler:** Tích hợp nút Restart dịch vụ Print Spooler nhanh chóng mà không cần vào Services.msc.

### 📊 Quản lý & Tiện ích
5.  **Menu Chuột Phải Thông Minh:**
    *   ⭐ Đặt máy in mặc định (Set Default).
    *   ⚙️ Mở nhanh **Printing Preferences** & **Printer Properties** (Rất tiện lợi thay vì phải tìm trong Control Panel).
    *   🔄 Bật/Tắt chia sẻ mạng LAN (Toggle Sharing).
    *   📄 Xem chi tiết hàng đợi in (Queue).
    *   🖨️ In trang Test (Windows Test Page).
    *   🗑️ Xóa máy in.
6.  **Soi Lệnh In (Queue Viewer):** Xem danh sách file đang chờ in (Tên tài liệu, Người in, Số trang, Dung lượng KB/MB, Thời gian in).
7.  **Xuất Báo Cáo:** Xuất danh sách toàn bộ máy in ra file Excel (`.csv`) bao gồm: Tên, Cổng (Port), Driver, Trạng thái chia sẻ.
8.  **Công Cụ Mạng:** Tự động tách IP từ cổng máy in và Ping kiểm tra kết nối (Online/Offline) ngay trên menu chuột phải.

### 🪟 Tích hợp Hệ thống
9.  **Lối tắt tiện dụng:**
    *   ➕ **Thêm Máy In:** Mở nhanh giao diện Add Printer của Windows Settings.
    *   Mở nhanh *Print Management (MSC)* và *Devices & Printers*.
10. **Nhật ký (Log):** Tự động lưu lịch sử thao tác vào file `activity.log` để tra cứu lỗi.
11. **Giao diện tối ưu:** Đã cập nhật lại thứ tự cột hiển thị (Port đứng trước Driver) giúp dễ quan sát IP máy in hơn.

## 💻 Yêu cầu hệ thống

*   **Hệ điều hành:** Windows 7, 10, 11 (32-bit & 64-bit).
*   **Quyền hạn:** Bắt buộc chạy bằng quyền **Administrator** (Do phần mềm can thiệp Service và Registry).

## 📝 Hướng dẫn sử dụng

### Bước 1: Khởi động
Nhấn chuột phải vào file `PrinterManager.exe` ➔ Chọn **Run as Administrator**.

### Bước 2: Các thao tác chính
*   **Quét danh sách:** Nhấn **🔄 Quét / Làm mới** để tải danh sách máy in từ Registry.
*   **Thao tác nhanh:** **Click chuột phải** vào dòng máy in bất kỳ để mở Menu chức năng (In test, Ping IP, Chia sẻ...).
*   **Dọn dẹp Driver:** Nhấn **🧹 Xóa Driver không sử dụng** ➔ Phần mềm sẽ liệt kê các driver thừa ➔ Chọn và xóa (Spooler sẽ tự khởi động lại).

### Bước 3: Quản lý nâng cao
*   **Xem ai đang in:** Chuột phải vào máy in ➔ Chọn **📄 Xem chi tiết Lệnh in**.
*   **Xuất Excel:** Nhấn **📥 Xuất Báo Cáo** để lưu file `.csv`.

---

## ⚠️ Lưu ý quan trọng

1.  **Backup:** Phần mềm tự động tạo file `.reg` sao lưu cấu hình máy in vào thư mục `Backup/` trước khi bạn thực hiện xóa máy in.
2.  **Antivirus:** Một số trình diệt virus có thể cảnh báo nhầm do hành vi can thiệp Registry/Stop Service. Vui lòng thêm vào danh sách loại trừ nếu cần thiết.

---

## 📞 Thông tin
*   **Tác giả:** @danhcp
*   **Phiên bản:** 2.2.1

---
*Cảm ơn bạn đã sử dụng Printer Manager!*