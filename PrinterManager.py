import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import winreg
import subprocess
import ctypes
import os
import sys
import threading
import datetime
import re
import socket
import csv
import shutil

# --- HÀM HỖ TRỢ ---
def resource_path(relative_path):
    try: base_path = sys._MEIPASS
    except: base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def is_admin():
    try: return ctypes.windll.shell32.IsUserAnAdmin()
    except: return False

def run_as_admin():
    try:
        if getattr(sys, 'frozen', False):
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, None, None, 1)
        else:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{os.path.abspath(sys.argv[0])}"', None, 1)
        return True
    except: return False

# --- CẤU HÌNH ---
APP_VERSION = "1.4.2"
APP_BUILD = "Column_Swapped"
APP_AUTHOR = "@danhcp"
APP_TITLE = f"Printer Manager"
ICON_NAME = "printer.ico"

REG_PRINTERS = r"SYSTEM\CurrentControlSet\Control\Print\Printers"
REG_DRIVERS_V3 = r"SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3"
REG_DRIVERS_V4 = r"SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-4"
SPOOL_DIR = r"C:\Windows\System32\spool\PRINTERS"
BACKUP_DIR = "Backup"
LOG_FILE = "activity.log"

PRINTER_ATTRIBUTE_SHARED = 0x00000008

class CleanPrinterApp:
    def __init__(self, root):
        self.root = root
        
        # Setup Icon & ID
        try: ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(f'danhcp.printermanager.{APP_BUILD}')
        except: pass
        try: self.root.iconbitmap(resource_path(ICON_NAME))
        except: pass

        self.root.title(f"{APP_TITLE} v{APP_VERSION} | Chế Độ Admin")
        self.root.geometry("1000x670") 
        self.root.resizable(False, False)
        
        self.setup_ui()
        self.create_context_menu()
        
        self.log(f"Khởi động {APP_TITLE} - Ver {APP_VERSION} by @danhcp")
        
        # Kiểm tra Spooler ngay khi mở
        self.check_spooler_status_on_startup()
        
        # Hướng dẫn người dùng
        self.log("ℹ️ Vui lòng nhấn nút '🔄 Quét / Làm mới' để tải danh sách máy in.")

    def setup_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=0)

        main_container = ttk.Frame(self.root)
        main_container.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        main_container.columnconfigure(0, weight=1) 
        main_container.columnconfigure(1, weight=0)
        main_container.rowconfigure(0, weight=1)

        # === CỘT TRÁI: DANH SÁCH ===
        frame_list = ttk.LabelFrame(main_container, text=" 🖨️ Danh Sách Máy In (Click chuột phải để xem menu) ", padding=5)
        frame_list.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        # --- CẬP NHẬT CỘT: Đổi vị trí Port và Driver ---
        columns = ("no", "status", "name", "port", "driver", "share") 
        self.tree = ttk.Treeview(frame_list, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("no", text="STT")
        self.tree.heading("status", text="Trạng thái") 
        self.tree.heading("name", text="Tên Máy In")
        self.tree.heading("port", text="Cổng (Port)") # Đổi lên trước
        self.tree.heading("driver", text="Driver")     # Đổi xuống sau
        self.tree.heading("share", text="Chia sẻ")
        
        self.tree.column("no", width=40, anchor="center")
        self.tree.column("status", width=100, anchor="center") 
        self.tree.column("name", width=220)
        self.tree.column("port", width=120)   # Đổi lên trước
        self.tree.column("driver", width=180) # Đổi xuống sau
        self.tree.column("share", width=80, anchor="center")

        # Scrollbar Fix
        v_scrollbar = ttk.Scrollbar(frame_list, orient="vertical", command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(frame_list, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        frame_list.grid_rowconfigure(0, weight=1)
        frame_list.grid_columnconfigure(0, weight=1)

        self.tree.bind("<Button-3>", self.show_context_menu)

        # Tag màu
        self.tree.tag_configure('default_printer', foreground='#c0392b', font=('Segoe UI', 9, 'bold'))
        self.tree.tag_configure('offline', foreground='#7f8c8d') 

        # === CỘT PHẢI: CHỨC NĂNG ===
        frame_controls = ttk.Frame(main_container, width=260)
        frame_controls.grid(row=0, column=1, sticky="ns")
        frame_controls.pack_propagate(False)

        grp_pad_y = 5
        btn_pad_y = 3

        # Box 1: Thao tác chính
        gb_main = ttk.LabelFrame(frame_controls, text=" ⚡ Thao Tác Chính ", padding=5)
        gb_main.pack(fill="x", pady=(0, grp_pad_y))
        ttk.Button(gb_main, text="🔄 Quét / Làm mới", command=self.scan_printers).pack(fill="x", pady=btn_pad_y)
        ttk.Button(gb_main, text="➕ Thêm Máy In", command=self.action_add_printer).pack(fill="x", pady=btn_pad_y)
        ttk.Button(gb_main, text="🧹 Xóa Driver không sử dụng", command=self.action_delete_unused_drivers).pack(fill="x", pady=btn_pad_y)

        # Box 2: Bảo trì & Khắc phục sự cố
        gb_maint = ttk.LabelFrame(frame_controls, text=" 🛠️ Bảo trì & Khắc phục sự cố ", padding=5)
        gb_maint.pack(fill="x", pady=grp_pad_y)
        ttk.Button(gb_maint, text="♻️ Restart Spooler", command=lambda: self.run_thread(self.restart_spooler)).pack(fill="x", pady=btn_pad_y)
        ttk.Button(gb_maint, text="🧹 Xóa Lệnh Kẹt (Clear)", command=self.clear_spool_files).pack(fill="x", pady=btn_pad_y)

        # Box 3: Báo cáo & Công cụ Windows
        gb_reports_win = ttk.LabelFrame(frame_controls, text=" 📊 Báo cáo & Công cụ Windows ", padding=5)
        gb_reports_win.pack(fill="x", pady=grp_pad_y)
        ttk.Button(gb_reports_win, text="📥 Xuất Báo Cáo (Excel)", command=self.export_report).pack(fill="x", pady=btn_pad_y)
        ttk.Button(gb_reports_win, text="📂 Printer Management (MSC)", command=lambda: self.run_cmd("printmanagement.msc")).pack(fill="x", pady=btn_pad_y)
        ttk.Button(gb_reports_win, text="⚙️ Devices & Printers (Control)", command=lambda: self.run_cmd("control printers")).pack(fill="x", pady=btn_pad_y)

        # === NHẬT KÝ ===
        frame_log = ttk.LabelFrame(main_container, text=" 📟 Nhật Ký Hoạt Động ", padding=2)
        frame_log.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(5, 0))
        main_container.rowconfigure(1, weight=0)
        self.txt_log = scrolledtext.ScrolledText(frame_log, height=6, state='disabled', font=("Consolas", 8)) 
        self.txt_log.pack(fill="both", expand=True)

        # === FOOTER ===
        frame_footer = ttk.Frame(self.root, padding=(5, 2))
        frame_footer.grid(row=1, column=0, sticky="ew")
        ttk.Label(frame_footer, text=f"Phiên bản {APP_VERSION}", font=("Segoe UI", 8), foreground="#555").pack(side="left")
        ttk.Label(frame_footer, text=f"Tác giả: {APP_AUTHOR}", font=("Segoe UI", 8, "bold"), foreground="#0055aa").pack(side="right")

    # --- HỆ THỐNG LOG ---
    def log(self, msg):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_msg = f"[{timestamp}] {msg}"
        try:
            self.txt_log.config(state='normal')
            self.txt_log.insert("end", f"> {msg}\n")
            self.txt_log.see("end")
            self.txt_log.config(state='disabled')
        except: pass
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(log_msg + "\n")
        except: pass

    def run_cmd(self, cmd):
        try: subprocess.Popen(cmd, shell=True)
        except Exception as e: self.log(f"Lỗi mở lệnh: {e}")

    def run_thread(self, func, args=()):
        threading.Thread(target=func, args=args, daemon=True).start()

    def check_spooler_status_on_startup(self):
        # Hàm kiểm tra trạng thái spooler
        def _check():
            try:
                cmd = 'powershell "Get-Service spooler | Select-Object -ExpandProperty Status"'
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                out = subprocess.check_output(cmd, shell=True, startupinfo=si).decode().strip()
                
                status_icon = "🟢" if out == "Running" else "🔴"
                msg = f"Print Spooler Service: {status_icon} {out}"
                self.root.after(0, lambda: self.log(msg))
            except Exception as e:
                self.root.after(0, lambda: self.log(f"Không thể kiểm tra Spooler: {e}"))
        self.run_thread(_check)

    def get_default_printer_name(self):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows NT\CurrentVersion\Windows") as key:
                device_string, _ = winreg.QueryValueEx(key, "Device")
                return device_string.split(',')[0]
        except: return None

    def get_printer_statuses_map(self):
        status_map = {}
        try:
            cmd = 'powershell "Get-Printer | Select-Object Name, PrinterStatus | ConvertTo-Csv -NoTypeInformation"'
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, text=True, startupinfo=si)
            out, _ = process.communicate()
            
            lines = out.strip().splitlines()
            if len(lines) > 1:
                reader = csv.reader(lines)
                next(reader) 
                for row in reader:
                    if len(row) >= 2:
                        status_map[row[0]] = row[1]
        except: pass
        return status_map

    def translate_status(self, status_str):
        s = status_str.lower()
        if s == "normal" or s == "idle": return "🟢 Sẵn sàng"
        if s == "printing": return "🖨️ Đang in"
        if s == "paused": return "⏸️ Tạm dừng"
        if s == "error": return "🔴 Lỗi"
        if s == "offline": return "⚫ Offline"
        if s == "paperjam": return "⚠️ Kẹt giấy"
        if s == "dooropen": return "⚠️ Nắp mở"
        return f"⚪ {status_str}"

    # --- MENU CHUỘT PHẢI ---
    def create_context_menu(self):
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="⭐ Đặt làm máy in Mặc định", command=self.set_default_printer)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="⚙️ Printing Preferences...", command=self.open_printing_preferences)
        self.context_menu.add_command(label="🔧 Printer Properties...", command=self.open_printer_properties)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="📄 Xem chi tiết Lệnh in (Queue)", command=self.view_print_queue)
        self.context_menu.add_command(label="🖨️ In thử (Test Page)", command=self.action_print_test)
        self.context_menu.add_command(label="🌐 Ping IP Máy in", command=lambda: self.run_thread(self.action_ping))
        self.context_menu.add_separator()
        self.context_menu.add_command(label="🔄 Bật/Tắt Chia sẻ LAN", command=self.toggle_sharing)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="🗑️ Xóa Máy In...", command=self.action_delete)

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    # --- CÁC TÍNH NĂNG ---
    def open_printing_preferences(self):
        sel = self.tree.selection()
        if not sel: return
        try:
            raw_name = self.tree.item(sel[0])['values'][2] 
            p_name = raw_name.lstrip("⭐").lstrip()
            cmd = f'rundll32 printui.dll,PrintUIEntry /e /n "{p_name}"'
            self.log(f"Mở Printing Preferences cho: {p_name}")
            self.run_cmd(cmd)
        except Exception as e: self.log(f"Lỗi: {e}")

    def open_printer_properties(self):
        sel = self.tree.selection()
        if not sel: return
        try:
            raw_name = self.tree.item(sel[0])['values'][2]
            p_name = raw_name.lstrip("⭐").lstrip()
            cmd = f'rundll32 printui.dll,PrintUIEntry /p /n "{p_name}"'
            self.log(f"Mở Printer Properties cho: {p_name}")
            self.run_cmd(cmd)
        except Exception as e: self.log(f"Lỗi: {e}")

    def set_default_printer(self):
        sel = self.tree.selection()
        if not sel: return
        raw_name = self.tree.item(sel[0])['values'][2]
        p_name = raw_name.lstrip("⭐").lstrip()
        self.log(f"⭐ Đang đặt '{p_name}' làm mặc định...")
        try:
            subprocess.run(f'rundll32 printui.dll,PrintUIEntry /y /n "{p_name}"', shell=True)
            messagebox.showinfo("Thành công", f"Đã đặt '{p_name}' làm máy in mặc định.")
            self.scan_printers()
        except Exception as e: self.log(f"❌ Lỗi: {e}")

    def toggle_sharing(self):
        sel = self.tree.selection()
        if not sel: return
        raw_name = self.tree.item(sel[0])['values'][2]
        p_name = raw_name.lstrip("⭐").lstrip()
        try:
            cmd_check = f'powershell "Get-Printer -Name \'{p_name}\' | Select-Object -ExpandProperty Shared"'
            out = subprocess.check_output(cmd_check, shell=True).decode().strip().lower()
            is_shared = (out == 'true')
            new_status = not is_shared
            action_text = "BẬT Chia sẻ" if new_status else "TẮT Chia sẻ"
            
            if messagebox.askyesno("Chia sẻ", f"Máy in: {p_name}\nBạn muốn {action_text}?"):
                ps_bool = "$true" if new_status else "$false"
                cmd_set = f'powershell "Set-Printer -Name \'{p_name}\' -Shared {ps_bool}"'
                subprocess.run(cmd_set, shell=True)
                self.log(f"🔄 Đã {action_text} cho {p_name}")
                self.scan_printers() 
        except: messagebox.showerror("Lỗi", "Không thể thay đổi chia sẻ.")

    def view_print_queue(self):
        sel = self.tree.selection()
        if not sel: return
        raw_name = self.tree.item(sel[0])['values'][2]
        p_name = raw_name.lstrip("⭐").lstrip()
        
        top = tk.Toplevel(self.root)
        top.title(f"Lệnh in: {p_name}")
        top.geometry("750x400")
        cols = ("id", "doc", "user", "pages", "size", "time")
        tree = ttk.Treeview(top, columns=cols, show="headings")
        tree.heading("id", text="ID"); tree.column("id", width=40)
        tree.heading("doc", text="Tên Tài Liệu"); tree.column("doc", width=200)
        tree.heading("user", text="Người in"); tree.column("user", width=100)
        tree.heading("pages", text="Trang"); tree.column("pages", width=50)
        tree.heading("size", text="Size"); tree.column("size", width=80)
        tree.heading("time", text="Thời gian"); tree.column("time", width=130)
        tree.pack(fill="both", expand=True)

        self.log(f"📄 Soi lệnh in: {p_name}...")
        def fetch_jobs():
            cmd = f'powershell "Get-PrintJob -PrinterName \'{p_name}\' | Select-Object Id,DocumentName,UserName,TotalPages,JobSize,SubmittedTime | ConvertTo-Csv -NoTypeInformation"'
            try:
                si = subprocess.STARTUPINFO(); si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                process = subprocess.Popen(cmd, stdout=subprocess.PIPE, text=True, startupinfo=si)
                out, _ = process.communicate()
                if not top.winfo_exists(): return 
                lines = out.strip().splitlines()
                if len(lines) > 1:
                    reader = csv.reader(lines)
                    next(reader)
                    for row in reader:
                        if len(row) >= 6:
                            try: size_kb = f"{int(row[4])/1024:.1f} KB"
                            except: size_kb = row[4]
                            tree.insert("", "end", values=(row[0], row[1], row[2], row[3], size_kb, row[5]))
                else: tree.insert("", "end", values=("Trống", "Không có tài liệu nào", "", "", "", ""))
            except Exception as e:
                try: tree.insert("", "end", values=("Lỗi", str(e), "", "", "", ""))
                except: pass
        self.run_thread(fetch_jobs)

    def export_report(self):
        filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV File", "*.csv")], title="Lưu Báo Cáo")
        if not filename: return
        self.log("📥 Đang xuất báo cáo...")
        try:
            with open(filename, mode='w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                # Cập nhật Header CSV theo thứ tự cột mới
                writer.writerow(["STT", "Trang Thai", "Ten May In", "Cong Ket Noi", "Driver", "Chia Se"])
                for item in self.tree.get_children():
                    row = self.tree.item(item)['values']
                    writer.writerow(row)
            messagebox.showinfo("Thành công", f"Đã lưu file:\n{filename}")
            self.log(f"✅ Xuất báo cáo OK.")
        except Exception as e: messagebox.showerror("Lỗi", str(e))

    def scan_printers(self):
        # Chạy trên thread riêng để không đơ UI
        self.run_thread(self._scan_printers_worker)

    def _scan_printers_worker(self):
        self.root.after(0, lambda: [self.tree.delete(item) for item in self.tree.get_children()])
        self.log("⏳ Đang quét danh sách & trạng thái...")
        
        status_map = self.get_printer_statuses_map()
        default_printer = self.get_default_printer_name()
        
        try:
            hKey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, REG_PRINTERS)
            idx = 0; count = 0
            items_to_insert = []

            while True:
                try:
                    p_name = winreg.EnumKey(hKey, idx)
                    
                    tags_to_apply = []
                    display_name = p_name
                    if p_name == default_printer:
                        display_name = f"⭐ {p_name}"
                        tags_to_apply.append('default_printer')

                    raw_status = status_map.get(p_name, "Unknown")
                    display_status = self.translate_status(raw_status)
                    if "Offline" in display_status:
                        tags_to_apply.append('offline')

                    d_name = "N/A"; port_name = "N/A"; share_status = "-"
                    try:
                        sub = winreg.OpenKey(hKey, p_name)
                        try: d_name, _ = winreg.QueryValueEx(sub, "Printer Driver")
                        except: pass
                        try: port_name, _ = winreg.QueryValueEx(sub, "Port")
                        except: pass
                        try: 
                            attr, _ = winreg.QueryValueEx(sub, "Attributes")
                            if attr & PRINTER_ATTRIBUTE_SHARED: share_status = "✅ CÓ"
                        except: pass
                        winreg.CloseKey(sub)
                    except: pass
                    
                    count += 1
                    # CẬP NHẬT THỨ TỰ DATA: Port đứng trước Driver
                    # 0: STT, 1: Status, 2: Name, 3: Port, 4: Driver, 5: Share
                    items_to_insert.append({
                        'values': (count, display_status, display_name, port_name, d_name, share_status),
                        'tags': tuple(tags_to_apply)
                    })
                    idx += 1
                except OSError: break
            winreg.CloseKey(hKey)
            
            def update_ui():
                for item in items_to_insert:
                    self.tree.insert("", "end", values=item['values'], tags=item['tags'])
                self.log(f"✅ Tìm thấy {count} máy in.")
            
            self.root.after(0, update_ui)

        except Exception as e: 
            self.log(f"❌ Lỗi Registry: {e}")

    def action_add_printer(self):
        try:
            self.run_cmd('start ms-settings:printers')
            messagebox.showinfo("Thông báo", "Đã mở cài đặt máy in.")
        except: pass

    def action_delete(self):
        sel = self.tree.selection()
        if not sel: 
            messagebox.showwarning("Chú ý", "Hãy chọn một máy in để xóa!")
            return
        data = self.tree.item(sel[0])['values']
        raw_name = data[2] # Index 2: Name
        p_name = raw_name.lstrip("⭐").lstrip()
        # Cập nhật Index: Driver bây giờ nằm ở vị trí 4
        d_name = data[4] 
        if messagebox.askyesno("Xác nhận Xóa", f"Xóa: {p_name}?\n\n(Tự động Backup trước khi xóa)"):
            self.run_thread(self.process_delete, args=(p_name, d_name))

    def action_delete_unused_drivers(self):
        self.log("🧹 Bắt đầu quét driver không sử dụng...")
        self.run_thread(self._find_and_show_unused_drivers)

    def _find_and_show_unused_drivers(self):
        try:
            in_use_drivers = set()
            hKey_printers = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, REG_PRINTERS)
            idx = 0
            while True:
                try:
                    sub_key = winreg.OpenKey(hKey_printers, winreg.EnumKey(hKey_printers, idx))
                    try:
                        dn, _ = winreg.QueryValueEx(sub_key, "Printer Driver")
                        if dn != "N/A": in_use_drivers.add(dn)
                    except: pass
                    winreg.CloseKey(sub_key)
                    idx += 1
                except: break
            winreg.CloseKey(hKey_printers)

            all_drivers = set()
            for path in [REG_DRIVERS_V3, REG_DRIVERS_V4]:
                try:
                    hKey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path)
                    idx = 0
                    while True:
                        try: all_drivers.add(winreg.EnumKey(hKey, idx)); idx+=1
                        except: break
                    winreg.CloseKey(hKey)
                except: pass

            unused = sorted(list(all_drivers - in_use_drivers))
            self.root.after(0, self._show_unused_driver_dialog, unused)
        except Exception as e: self.log(f"Lỗi quét driver: {e}")

    def _show_unused_driver_dialog(self, unused_drivers):
        if not unused_drivers:
            messagebox.showinfo("Hoàn tất", "Hệ thống sạch, không có driver thừa.")
            return
        
        top = tk.Toplevel(self.root)
        top.title("Xóa Driver Rác")
        top.geometry("600x400")
        ttk.Label(top, text=f"Tìm thấy {len(unused_drivers)} driver không dùng:").pack(pady=5)
        
        frame = ttk.Frame(top)
        frame.pack(fill="both", expand=True, padx=10)
        lb = tk.Listbox(frame, selectmode="extended")
        for d in unused_drivers: lb.insert("end", d)
        lb.pack(side="left", fill="both", expand=True)
        
        def do_del():
            sels = [lb.get(i) for i in lb.curselection()]
            if not sels: return
            if messagebox.askyesno("Xóa", f"Xóa {len(sels)} driver? Spooler sẽ restart."):
                self.run_thread(self._process_delete_drivers, args=(sels,))
                top.destroy()
        
        ttk.Button(top, text="Xóa Đã Chọn", command=do_del).pack(pady=5)

    def _process_delete_drivers(self, drivers):
        self.stop_spooler()
        for d in drivers: self.delete_driver_reg(d)
        self.start_spooler()
        self.log(f"Đã xóa {len(drivers)} driver.")
        self.root.after(500, self.scan_printers)

    def process_delete(self, p_name, d_name):
        self.log(f"--- XÓA: {p_name} ---")
        self.backup_registry(p_name)
        self.stop_spooler()
        path = f"{REG_PRINTERS}\\{p_name}"
        if self.delete_registry_tree(winreg.HKEY_LOCAL_MACHINE, path): self.log("✅ Xóa Registry OK.")
        if d_name and d_name != "N/A": self.delete_driver_reg(d_name)
        self.start_spooler()
        self.log("--- HOÀN TẤT ---")
        self.root.after(1000, self.scan_printers)

    def backup_registry(self, p_name):
        try:
            if not os.path.exists(BACKUP_DIR): os.makedirs(BACKUP_DIR)
            safe = "".join([c for c in p_name if c.isalnum() or c in (' ','-','_')]).strip()
            fname = os.path.join(BACKUP_DIR, f"{safe}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.reg")
            subprocess.run(f'reg export "HKLM\\{REG_PRINTERS}\\{p_name}" "{fname}" /y', shell=True, creationflags=0x08000000)
            self.log(f"💾 Backup: {fname}")
        except: pass

    def delete_registry_tree(self, root, path):
        try:
            open_key = winreg.OpenKey(root, path, 0, winreg.KEY_ALL_ACCESS)
            while True:
                try: sub = winreg.EnumKey(open_key, 0); self.delete_registry_tree(root, f"{path}\\{sub}")
                except: break
            winreg.CloseKey(open_key); winreg.DeleteKey(root, path); return True
        except: return False

    def delete_driver_reg(self, d_name):
        paths = [REG_DRIVERS_V3, REG_DRIVERS_V4]
        for base in paths:
            full = f"{base}\\{d_name}"
            try: 
                winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, full)
                if self.delete_registry_tree(winreg.HKEY_LOCAL_MACHINE, full): self.log(f"✅ Xóa Driver: {d_name}")
            except: pass

    def stop_spooler(self): subprocess.run("net stop spooler", shell=True, creationflags=0x08000000)
    def start_spooler(self): subprocess.run("net start spooler", shell=True, creationflags=0x08000000)
    def restart_spooler(self): self.stop_spooler(); self.start_spooler(); self.log("✅ Spooler Restarted.")

    def clear_spool_files(self):
        if not messagebox.askyesno("Confirm", "Xóa hết lệnh in đang chờ?"): return
        self.stop_spooler()
        try:
            if os.path.exists(SPOOL_DIR):
                for f in os.listdir(SPOOL_DIR):
                    fp = os.path.join(SPOOL_DIR, f)
                    try:
                        if os.path.isfile(fp): os.unlink(fp)
                        elif os.path.isdir(fp): shutil.rmtree(fp)
                    except: pass
            self.log("✅ Dọn sạch Queue.")
        except: pass
        self.start_spooler()

    def action_print_test(self):
        try:
            sel = self.tree.selection()
            if sel:
                raw_name = self.tree.item(sel[0])['values'][2] # Index 2: Name
                p_name = raw_name.lstrip("⭐").lstrip()
                subprocess.Popen(f'rundll32 printui.dll,PrintUIEntry /k /n "{p_name}"', shell=True)
                self.log(f"🖨️ In test: {p_name}")
        except: pass

    def action_ping(self):
        try:
            sel = self.tree.selection()
            if sel:
                # Cập nhật Index: Port bây giờ nằm ở vị trí 3 (Index 3)
                port = str(self.tree.item(sel[0])['values'][3]) 
                ip = re.search(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', port)
                if ip:
                    ip = ip.group(1)
                    self.log(f"🌐 Ping {ip}...")
                    si = subprocess.STARTUPINFO(); si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    proc = subprocess.Popen(f"ping -n 4 {ip}", stdout=subprocess.PIPE, startupinfo=si)
                    out, _ = proc.communicate()
                    if b"TTL=" in out: messagebox.showinfo("OK", f"✅ {ip} Online")
                    else: messagebox.showerror("Fail", f"❌ {ip} Offline")
                else: messagebox.showwarning("Thông báo", "Không tìm thấy IP.")
        except: pass

if __name__ == "__main__":
    if is_admin():
        root = tk.Tk()
        app = CleanPrinterApp(root)
        root.mainloop()
    else:
        if run_as_admin(): sys.exit()