import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import winreg
import subprocess
import ctypes
import os
import sys
import threading

# --- HÀM HỖ TRỢ TÌM FILE KHI ĐÓNG GÓI EXE ---
def resource_path(relative_path):
    """ Lấy đường dẫn tuyệt đối tới tài nguyên, dùng cho cả Dev và PyInstaller """
    try:
        # PyInstaller tạo ra thư mục tạm _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- THÔNG TIN ỨNG DỤNG ---
APP_VERSION = "1.0"
APP_BUILD = "091225"
APP_AUTHOR = "@danhcp"
APP_TITLE = f"Quản Lý Registry Máy In"
ICON_NAME = "printer.ico"  # Tên file icon của bạn

# --- CẤU HÌNH REGISTRY ---
REG_PRINTERS = r"SYSTEM\CurrentControlSet\Control\Print\Printers"
REG_DRIVERS_V3 = r"SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3"
REG_DRIVERS_V4 = r"SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-4"
SPOOL_DIR = r"C:\Windows\System32\spool\PRINTERS"

class CleanPrinterApp:
    def __init__(self, root):
        self.root = root
        
        # --- CẤU HÌNH ICON & TASKBAR ID ---
        # 1. Fix lỗi icon dưới Taskbar (Windows nhóm các cửa sổ python lại, cần tách ID ra)
        try:
            myappid = f'danhcp.printermanager.v1.{APP_BUILD}' # ID tùy ý
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except:
            pass

        # 2. Set Icon cho cửa sổ và Taskbar
        icon_path = resource_path(ICON_NAME)
        try:
            self.root.iconbitmap(icon_path)
        except Exception:
            # Nếu không thấy icon thì bỏ qua (dùng mặc định)
            pass

        # 3. Tiêu đề & Kích thước
        self.root.title(f"{APP_TITLE} (Build {APP_BUILD}) | Chế độ Admin")
        self.root.geometry("850x500")
        self.root.resizable(False, False)
        
        # Kiểm tra quyền Admin
        if not self.is_admin():
            messagebox.showerror("⚠️ Lỗi Quyền", "Vui lòng chạy phần mềm bằng quyền Administrator (Run as Administrator).")
            root.destroy()
            return

        self.setup_ui()
        self.log(f"Đã khởi động {APP_TITLE} - Ver {APP_VERSION} (Build {APP_BUILD})")
        self.log(f"Tác giả: {APP_AUTHOR}")

    def is_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def setup_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=0)

        # --- KHUNG CHÍNH ---
        main_container = ttk.Frame(self.root)
        main_container.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        main_container.columnconfigure(0, weight=1) 
        main_container.columnconfigure(1, weight=0)
        main_container.rowconfigure(0, weight=1)

        # === 1. CỘT TRÁI: DANH SÁCH MÁY IN ===
        frame_list = ttk.LabelFrame(main_container, text=" Danh Sách Máy In (Registry) ", padding=5)
        frame_list.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        columns = ("no", "name", "driver", "port")
        self.tree = ttk.Treeview(frame_list, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("no", text="STT")
        self.tree.heading("name", text="Tên Máy In")
        self.tree.heading("driver", text="Driver Đang Dùng")
        self.tree.heading("port", text="Cổng (Port)")
        
        self.tree.column("no", width=35, anchor="center", stretch=False)
        self.tree.column("name", width=220)
        self.tree.column("driver", width=180)
        self.tree.column("port", width=100)

        scrollbar = ttk.Scrollbar(frame_list, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # === 2. CỘT PHẢI: CHỨC NĂNG ===
        frame_controls = ttk.Frame(main_container, width=220)
        frame_controls.grid(row=0, column=1, sticky="ns")
        frame_controls.pack_propagate(False)

        btn_pad_y = 2
        grp_pad_y = 5

        # [Box 1] Thao tác chính
        gb_main = ttk.LabelFrame(frame_controls, text=" Thao Tác Chính ", padding=5)
        gb_main.pack(fill="x", pady=(0, grp_pad_y))

        ttk.Button(gb_main, text="🔄 Quét / Làm mới", command=self.scan_printers).pack(fill="x", pady=btn_pad_y)
        ttk.Separator(gb_main, orient="horizontal").pack(fill="x", pady=4)
        ttk.Button(gb_main, text="🗑️ XÓA MÁY IN ĐANG CHỌN", command=self.action_delete).pack(fill="x", pady=btn_pad_y)

        # [Box 2] Dịch vụ In
        gb_spool = ttk.LabelFrame(frame_controls, text=" Dịch Vụ In (Spooler) ", padding=5)
        gb_spool.pack(fill="x", pady=grp_pad_y)

        ttk.Button(gb_spool, text="♻️ Khởi động lại Spooler", command=lambda: self.run_thread(self.restart_spooler)).pack(fill="x", pady=btn_pad_y)
        ttk.Button(gb_spool, text="🧹 Xóa lệnh in bị kẹt", command=self.clear_spool_files).pack(fill="x", pady=btn_pad_y)

        # [Box 3] Công cụ Windows
        gb_tools = ttk.LabelFrame(frame_controls, text=" Công Cụ Windows ", padding=5)
        gb_tools.pack(fill="x", pady=grp_pad_y)
        
        ttk.Button(gb_tools, text="📂 Quản lý Máy in (MSC)", command=lambda: self.run_cmd("printmanagement.msc")).pack(fill="x", pady=btn_pad_y)
        ttk.Button(gb_tools, text="📝 Xem tất cả Driver (PnP)", command=self.scan_pnputil).pack(fill="x", pady=btn_pad_y)

        # === 3. NHẬT KÝ ===
        frame_log = ttk.LabelFrame(main_container, text=" Nhật Ký Hoạt Động ", padding=2)
        frame_log.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(5, 0))
        main_container.rowconfigure(1, weight=0)

        self.txt_log = scrolledtext.ScrolledText(frame_log, height=5, state='disabled', font=("Consolas", 8)) 
        self.txt_log.pack(fill="both", expand=True)

        # === 4. CHÂN TRANG (FOOTER) ===
        frame_footer = ttk.Frame(self.root, padding=(5, 2))
        frame_footer.grid(row=1, column=0, sticky="ew")

        lbl_ver = ttk.Label(frame_footer, text=f"Phiên bản {APP_VERSION} (Build {APP_BUILD})", font=("Segoe UI", 8), foreground="#555")
        lbl_ver.pack(side="left")

        lbl_auth = ttk.Label(frame_footer, text=f"Phát triển bởi: {APP_AUTHOR}", font=("Segoe UI", 8, "bold"), foreground="#0055aa")
        lbl_auth.pack(side="right")

    # --- LOGIC CODE ---
    def log(self, msg):
        self.txt_log.config(state='normal')
        self.txt_log.insert("end", f"> {msg}\n")
        self.txt_log.see("end")
        self.txt_log.config(state='disabled')

    def run_cmd(self, cmd):
        try:
            subprocess.Popen(cmd, shell=True)
        except Exception as e:
            self.log(f"❌ Lỗi thực thi lệnh: {e}")

    def run_thread(self, func):
        threading.Thread(target=func, daemon=True).start()

    def scan_printers(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.log("⏳ Đang quét dữ liệu Registry...")
        try:
            hKey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, REG_PRINTERS)
            idx = 0
            count = 0
            while True:
                try:
                    p_name = winreg.EnumKey(hKey, idx)
                    d_name = "Không xác định"
                    port_name = "Không xác định"
                    try:
                        sub = winreg.OpenKey(hKey, p_name)
                        try: d_name, _ = winreg.QueryValueEx(sub, "Printer Driver")
                        except: pass
                        try: port_name, _ = winreg.QueryValueEx(sub, "Port")
                        except: pass
                        winreg.CloseKey(sub)
                    except: pass
                    count += 1
                    self.tree.insert("", "end", values=(count, p_name, d_name, port_name))
                    idx += 1
                except OSError: break
            winreg.CloseKey(hKey)
            self.log(f"✅ Tìm thấy {count} máy in trong hệ thống.")
        except Exception as e:
            self.log(f"❌ Lỗi đọc Registry: {e}")

    def action_delete(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("⚠️ Chú ý", "Vui lòng chọn một máy in trong danh sách để xóa.")
            return
        data = self.tree.item(sel[0])['values']
        p_name = data[1]
        d_name = data[2]
        if messagebox.askyesno("⁉️ Xác nhận xóa", f"Bạn chuẩn bị xóa:\n\n- Máy in: {p_name}\n- Driver: {d_name}\n\nThao tác này sẽ can thiệp vào Registry. Tiếp tục?"):
            self.run_thread(lambda: self.process_delete(sel[0], p_name, d_name))

    def process_delete(self, item_id, p_name, d_name):
        self.log(f"--- BẮT ĐẦU XÓA: {p_name} ---")
        self.stop_spooler()
        path = f"{REG_PRINTERS}\\{p_name}"
        if self.delete_registry_tree(winreg.HKEY_LOCAL_MACHINE, path):
            self.log(f"✅ Đã xóa Key Registry máy in.")
            self.tree.delete(item_id)
        else:
            self.log(f"⚠️ Không tìm thấy Key Registry (có thể đã bị xóa trước đó).")
        if d_name and d_name != "Không xác định":
            self.delete_driver(d_name)
        self.start_spooler()
        self.log("--- QUY TRÌNH HOÀN TẤT ---")
        messagebox.showinfo("✅ Thành công", "Đã xóa xong máy in và driver.\nVui lòng khởi động lại máy tính.")

    def delete_driver(self, d_name):
        paths = [REG_DRIVERS_V3, REG_DRIVERS_V4]
        found = False
        for base in paths:
            full_path = f"{base}\\{d_name}"
            try:
                winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, full_path)
                if self.delete_registry_tree(winreg.HKEY_LOCAL_MACHINE, full_path):
                    self.log(f"✅ Đã xóa Key Driver: {d_name}")
                    found = True
            except: pass
        if not found:
            self.log(f"ℹ️ Không tìm thấy Driver này trong Registry.")

    def delete_registry_tree(self, root, path):
        try:
            open_key = winreg.OpenKey(root, path, 0, winreg.KEY_ALL_ACCESS)
            while True:
                try:
                    sub = winreg.EnumKey(open_key, 0)
                    self.delete_registry_tree(root, f"{path}\\{sub}")
                except OSError: break
            winreg.CloseKey(open_key)
            winreg.DeleteKey(root, path)
            return True
        except: return False

    def stop_spooler(self):
        self.log("⏳ Đang dừng dịch vụ Spooler...")
        subprocess.run("net stop spooler", shell=True, creationflags=subprocess.CREATE_NO_WINDOW)

    def start_spooler(self):
        self.log("⏳ Đang bật lại dịch vụ Spooler...")
        subprocess.run("net start spooler", shell=True, creationflags=subprocess.CREATE_NO_WINDOW)

    def restart_spooler(self):
        self.stop_spooler()
        self.start_spooler()
        self.log("✅ Dịch vụ in đã được khởi động lại.")

    def clear_spool_files(self):
        if not messagebox.askyesno("⚠️ Xác nhận", "Hành động này sẽ xóa toàn bộ lệnh in đang chờ (Queue).\nBạn có chắc chắn không?"): return
        self.run_thread(self._clear_spool_logic)

    def _clear_spool_logic(self):
        self.stop_spooler()
        self.log("🧹 Đang dọn dẹp thư mục Spool...")
        try:
            if os.path.exists(SPOOL_DIR):
                for f in os.listdir(SPOOL_DIR):
                    fp = os.path.join(SPOOL_DIR, f)
                    try:
                        if os.path.isfile(fp): os.unlink(fp)
                        elif os.path.isdir(fp): shutil.rmtree(fp)
                    except: pass
            self.log("✅ Đã dọn sạch lệnh in kẹt.")
        except: pass
        self.start_spooler()
        messagebox.showinfo("✅ Xong", "Đã xử lý xong.")

    def scan_pnputil(self):
        top = tk.Toplevel(self.root)
        top.title("📝 Danh Sách Driver (Pnputil)")
        top.geometry("650x450")
        txt = scrolledtext.ScrolledText(top, font=("Consolas", 9))
        txt.pack(fill="both", expand=True)
        try:
            res = subprocess.check_output("pnputil /enum-drivers", shell=True, encoding='cp1258', errors='ignore')
            txt.insert("end", res)
        except:
            txt.insert("end", "❌ Không thể lấy danh sách driver.")

if __name__ == "__main__":
    root = tk.Tk()
    app = CleanPrinterApp(root)
    root.mainloop()