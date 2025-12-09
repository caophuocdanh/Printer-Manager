# 🖨️ Printer Registry Manager (Portable)

**Công cụ Sửa lỗi & Xóa Máy in Tận gốc trên Windows**  
*Phiên bản: 1.0 (Build 091225)*  
*Tác giả: @danhcp*

![Platform](https://img.shields.io/badge/Platform-Windows-0078D6.svg?style=flat-square)
![Type](https://img.shields.io/badge/Type-Portable_(EXE)-orange.svg?style=flat-square)
![License](https://img.shields.io/badge/License-Freeware-green.svg?style=flat-square)

## 📖 Giới thiệu

**Printer Registry Manager** là phần mềm dạng **Portable** (chạy ngay không cần cài đặt), được thiết kế để giải quyết các sự cố máy in phổ biến mà công cụ Windows mặc định không xử lý được.

Công cụ này đặc biệt hữu ích khi:
*   Máy in bị kẹt, không thể xóa (Remove device không có tác dụng).
*   Lỗi "Driver is in use" hoặc không thể cài lại driver mới.
*   Máy in bị kẹt lệnh in (Queue) không thể hủy.
*   Cần dọn dẹp sạch sẽ hệ thống máy in cũ.

## ✨ Tính năng nổi bật

1.  **🗑️ Xóa Máy In Tận Gốc:** Can thiệp sâu vào Registry để xóa bỏ máy in và Driver đi kèm (ngay cả khi bị lỗi).
2.  **🧹 Xóa Lệnh In Kẹt (Clear Queue):** Một cú click để xóa sạch toàn bộ lệnh in đang bị treo trong hệ thống.
3.  **♻️ Tự Động Xử Lý Spooler:** Tự động Tắt/Bật dịch vụ Print Spooler để đảm bảo quá trình xóa không bị lỗi "Access Denied".
4.  **🚀 Không Cần Cài Đặt:** Chỉ cần tải về 1 file `.exe` duy nhất và chạy.
5.  **🛠️ Công Cụ Hỗ Trợ:** Tích hợp nút mở nhanh *Print Management* và *Pnputil* để kiểm tra hệ thống.

## 💻 Yêu cầu hệ thống

*   **Hệ điều hành:** Windows 7, Windows 10, Windows 11 (32-bit & 64-bit).
*   **Quyền hạn:** Bắt buộc chạy bằng quyền **Administrator**.

## 📝 Hướng dẫn sử dụng

### Bước 1: Mở phần mềm
Do phần mềm can thiệp vào hệ thống (Registry & Services), bạn **BẮT BUỘC** phải mở như sau:
1.  Nhấn chuột phải vào file `PrinterManager.exe`.
2.  Chọn **Run as Administrator** (Chạy với tư cách quản trị viên).

### Bước 2: Quét danh sách
*   Tại giao diện chính, nhấn nút **🔄 Quét / Làm mới**.
*   Danh sách các máy in hiện có trong Registry sẽ hiện ra (kèm tên Driver và Cổng kết nối).

### Bước 3: Xử lý sự cố
*   **Để xóa máy in:** Chọn tên máy in trong danh sách ➔ Nhấn **🗑️ XÓA MÁY IN ĐANG CHỌN** ➔ Chọn *Yes* để xác nhận.
*   **Để sửa lỗi kẹt lệnh in:** Nhấn nút **🧹 Xóa lệnh in bị kẹt**.
*   **Để khởi động lại dịch vụ in:** Nhấn nút **♻️ Khởi động lại Spooler**.

### Bước 4: Hoàn tất
*   Sau khi thực hiện xong các thao tác, vui lòng **Khởi động lại máy tính (Restart)** để Windows cập nhật lại cấu hình sạch.

---

## ⚠️ Lưu ý quan trọng

1.  **An toàn dữ liệu:** Phần mềm thực hiện thay đổi trên Registry. Mặc dù đã được kiểm tra kỹ lưỡng, bạn nên cẩn trọng khi xóa các máy in hệ thống (như *Microsoft Print to PDF*, *Fax*...). Chỉ nên xóa các máy in vật lý bị lỗi.
2.  **Phần mềm chống virus:** Một số phần mềm diệt virus có thể cảnh báo nhầm do file `.exe` can thiệp vào Registry. Đây là hành vi bình thường của công cụ sửa lỗi hệ thống.

---

## 📞 Thông tin liên hệ & Hỗ trợ

Nếu bạn gặp vấn đề khi sử dụng hoặc muốn góp ý tính năng mới vui lòng tự làm vì đây là vibecode :))

*   **Tác giả:** @danhcp
*   **Phiên bản:** 1.0 Stable

---
*Cảm ơn bạn đã sử dụng Printer Registry Manager!*