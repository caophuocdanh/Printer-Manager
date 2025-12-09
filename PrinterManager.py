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
APP_VERSION = "1.3.2"
APP_BUILD = "Windows_Tools_Added"
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
        self.root.geometry("1050x650") # Tăng chiều cao xíu để chứa thêm nút
        self.root.resizable(False, False)
        
        self.setup_ui()
        self.create_context_menu()
        
        self.log(f"Khởi động {APP_TITLE} - Ver {APP_VERSION} by @danhcp")

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

        columns = ("no", "name", "driver", "port", "share")
        self.tree = ttk.Treeview(frame_list, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("no", text="STT")
        self.tree.heading("name", text="Tên Máy In")
        self.tree.heading("driver", text="Driver")
        self.tree.heading("port", text="Cổng (Port)")
        self.tree.heading("share", text="Chia sẻ")
        
        self.tree.column("no", width=40, anchor="center")
        self.tree.column("name", width=250)
        self.tree.column("driver", width=180)
        self.tree.column("port", width=120)
        self.tree.column("share", width=100, anchor="center")

        scrollbar = ttk.Scrollbar(frame_list, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.tree.bind("<Button-3>", self.show_context_menu)

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
        ttk.Separator(gb_main, orient="horizontal").pack(fill="x", pady=4)
        ttk.Button(gb_main, text="🗑️ XÓA (Tự động Backup)", command=self.action_delete).pack(fill="x", pady=btn_pad_y)

        # Box 2: Quản lý App
        gb_mgmt = ttk.LabelFrame(frame_controls, text=" 📊 Dữ Liệu & Báo Cáo ", padding=5)
        gb_mgmt.pack(fill="x", pady=grp_pad_y)
        ttk.Button(gb_mgmt, text="📄 Soi Lệnh In (Queue)", command=self.view_print_queue).pack(fill="x", pady=btn_pad_y)
        ttk.Button(gb_mgmt, text="📥 Xuất Báo Cáo (Excel)", command=self.export_report).pack(fill="x", pady=btn_pad_y)

        # Box 3: Công cụ Windows (MỚI)
        gb_win = ttk.LabelFrame(frame_controls, text=" 🔧 Công Cụ Windows ", padding=5)
        gb_win.pack(fill="x", pady=grp_pad_y)
        # Nút mở printmanagement.msc
        ttk.Button(gb_win, text="📂 Printer Management (MSC)", command=lambda: self.run_cmd("printmanagement.msc")).pack(fill="x", pady=btn_pad_y)
        # Nút mở Devices and Printers cũ (phòng khi bản Home không có MSC)
        ttk.Button(gb_win, text="⚙️ Devices & Printers (Control)", command=lambda: self.run_cmd("control printers")).pack(fill="x", pady=btn_pad_y)

        # Box 4: Tiện ích & Spooler
        gb_util = ttk.LabelFrame(frame_controls, text=" 🛠️ Dịch Vụ & Mạng ", padding=5)
        gb_util.pack(fill="x", pady=grp_pad_y)
        ttk.Button(gb_util, text="🖨️ In Trang Test", command=self.action_print_test).pack(fill="x", pady=btn_pad_y)
        ttk.Button(gb_util, text="🌐 Ping IP Máy in", command=lambda: self.run_thread(self.action_ping)).pack(fill="x", pady=btn_pad_y)
        ttk.Separator(gb_util, orient="horizontal").pack(fill="x", pady=4)
        ttk.Button(gb_util, text="♻️ Restart Spooler", command=lambda: self.run_thread(self.restart_spooler)).pack(fill="x", pady=btn_pad_y)
        ttk.Button(gb_util, text="🧹 Xóa Lệnh Kẹt (Clear)", command=self.clear_spool_files).pack(fill="x", pady=btn_pad_y)

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
        self.txt_log.config(state='normal')
        self.txt_log.insert("end", f"> {msg}\n")
        self.txt_log.see("end")
        self.txt_log.config(state='disabled')
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(log_msg + "\n")
        except: pass

    def run_cmd(self, cmd):
        # Hàm chạy lệnh cmd không chờ kết quả để tránh treo UI
        try: subprocess.Popen(cmd, shell=True)
        except Exception as e: self.log(f"Lỗi mở lệnh: {e}")

    def run_thread(self, func):
        threading.Thread(target=func, daemon=True).start()

    # --- MENU CHUỘT PHẢI ---
    def create_context_menu(self):
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="⭐ Đặt làm máy in Mặc định", command=self.set_default_printer)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="🔄 Bật/Tắt Chia sẻ LAN", command=self.toggle_sharing)
        self.context_menu.add_command(label="📄 Xem chi tiết Lệnh in", command=self.view_print_queue)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="🖨️ In thử (Test Page)", command=self.action_print_test)

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    # --- CÁC TÍNH NĂNG ---
    def set_default_printer(self):
        sel = self.tree.selection()
        if not sel: return
        p_name = self.tree.item(sel[0])['values'][1]
        self.log(f"⭐ Đang đặt '{p_name}' làm mặc định...")
        try:
            subprocess.run(f'rundll32 printui.dll,PrintUIEntry /y /n "{p_name}"', shell=True)
            messagebox.showinfo("Thành công", f"Đã đặt '{p_name}' làm máy in mặc định.")
            self.log("✅ Đặt mặc định thành công.")
        except Exception as e: self.log(f"❌ Lỗi: {e}")

    def toggle_sharing(self):
        sel = self.tree.selection()
        if not sel: return
        p_name = self.tree.item(sel[0])['values'][1]
        cmd_check = f'powershell "Get-Printer -Name \'{p_name}\' | Select-Object -ExpandProperty Shared"'
        try:
            out = subprocess.check_output(cmd_check, shell=True).decode().strip().lower()
            is_shared = (out == 'true')
            new_status = not is_shared
            action_text = "BẬT Chia sẻ" if new_status else "TẮT Chia sẻ"
            trang_thai_hien_tai = "ĐANG SHARE" if is_shared else "KHÔNG SHARE"
            if messagebox.askyesno("Chia sẻ", f"Máy in: {p_name}\nTrạng thái: {trang_thai_hien_tai}\n\nBạn muốn {action_text}?"):
                ps_bool = "$true" if new_status else "$false"
                cmd_set = f'powershell "Set-Printer -Name \'{p_name}\' -Shared {ps_bool}"'
                subprocess.run(cmd_set, shell=True)
                self.log(f"🔄 Đã {action_text} cho {p_name}")
                self.scan_printers() 
        except Exception as e:
            self.log(f"❌ Lỗi: {e}")
            messagebox.showerror("Lỗi", "Không thể thay đổi chia sẻ (Cần bật Firewall).")

    def view_print_queue(self):
        sel = self.tree.selection()
        if not sel: 
            messagebox.showwarning("⚠️", "Vui lòng chọn máy in trước!")
            return
        p_name = self.tree.item(sel[0])['values'][1]
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
                rows_to_insert = []
                if len(lines) > 1:
                    reader = csv.reader(lines)
                    next(reader)
                    for row in reader:
                        if len(row) >= 6:
                            try: size_kb = f"{int(row[4])/1024:.1f} KB"
                            except: size_kb = row[4]
                            rows_to_insert.append((row[0], row[1], row[2], row[3], size_kb, row[5]))
                try:
                    if rows_to_insert:
                        for r in rows_to_insert: tree.insert("", "end", values=r)
                    else: tree.insert("", "end", values=("Trống", "Không có tài liệu nào", "", "", "", ""))
                except tk.TclError: pass
            except Exception as e:
                try: 
                    if top.winfo_exists(): tree.insert("", "end", values=("Lỗi", str(e), "", "", "", ""))
                except: pass
        self.run_thread(fetch_jobs)

    def export_report(self):
        filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV File", "*.csv")], title="Lưu Báo Cáo")
        if not filename: return
        self.log("📥 Đang xuất báo cáo...")
        try:
            with open(filename, mode='w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(["STT", "Ten May In", "Driver", "Cong Ket Noi", "Chia Se"])
                for item in self.tree.get_children():
                    row = self.tree.item(item)['values']
                    writer.writerow(row)
            messagebox.showinfo("Thành công", f"Đã lưu file:\n{filename}")
            self.log(f"✅ Xuất báo cáo OK.")
        except Exception as e: messagebox.showerror("Lỗi", str(e))

    def scan_printers(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        self.log("⏳ Đang quét Registry...")
        try:
            hKey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, REG_PRINTERS)
            idx = 0; count = 0
            while True:
                try:
                    p_name = winreg.EnumKey(hKey, idx)
                    d_name = "N/A"; port_name = "N/A"; share_status = "Không"
                    try:
                        sub = winreg.OpenKey(hKey, p_name)
                        try: d_name, _ = winreg.QueryValueEx(sub, "Printer Driver")
                        except: pass
                        try: port_name, _ = winreg.QueryValueEx(sub, "Port")
                        except: pass
                        try: 
                            attr, _ = winreg.QueryValueEx(sub, "Attributes")
                            if attr & PRINTER_ATTRIBUTE_SHARED: share_status = "✅ CÓ"
                            else: share_status = "-"
                        except: pass
                        winreg.CloseKey(sub)
                    except: pass
                    count += 1
                    self.tree.insert("", "end", values=(count, p_name, d_name, port_name, share_status))
                    idx += 1
                except OSError: break
            winreg.CloseKey(hKey)
            self.log(f"✅ Tìm thấy {count} máy in.")
        except Exception as e: self.log(f"❌ Lỗi Registry: {e}")

    def action_delete(self):
        sel = self.tree.selection()
        if not sel: 
            messagebox.showwarning("Chú ý", "Hãy chọn một máy in để xóa!")
            return
        data = self.tree.item(sel[0])['values']
        p_name = data[1]; d_name = data[2]
        if messagebox.askyesno("Xác nhận Xóa", f"Xóa: {p_name}?\n\n(Tự động Backup trước khi xóa)"):
            self.run_thread(lambda: self.process_delete(p_name, d_name))

    def process_delete(self, p_name, d_name):
        self.log(f"--- XÓA: {p_name} ---")
        self.backup_registry(p_name)
        self.stop_spooler()
        path = f"{REG_PRINTERS}\\{p_name}"
        if self.delete_registry_tree(winreg.HKEY_LOCAL_MACHINE, path): self.log("✅ Xóa Registry OK.")
        if d_name and d_name != "N/A": self.delete_driver_reg(d_name)
        self.start_spooler()
        self.log("--- HOÀN TẤT ---")
        self.root.after(500, self.scan_printers)

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
                p_name = self.tree.item(sel[0])['values'][1]
                subprocess.Popen(f'rundll32 printui.dll,PrintUIEntry /k /n "{p_name}"', shell=True)
                self.log(f"🖨️ In test: {p_name}")
        except: pass

    def action_ping(self):
        try:
            sel = self.tree.selection()
            if sel:
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