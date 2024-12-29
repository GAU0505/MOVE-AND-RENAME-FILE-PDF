import os
import pandas as pd
from tkinter import Tk, filedialog, messagebox, StringVar, Text, Scrollbar, Frame, Label, Button, OptionMenu, ttk
from datetime import datetime

class FileRenameApp:
    def __init__(self, master):
        self.master = master
        self.master.title("ĐỔI TÊN FILE PDF - HỒ XUÂN ÁNH")
        self.master.geometry("900x700")
        self.master.resizable(False, False)
        self.master.config(bg="#42f5ec")

        # Biến trạng thái
        self.excel_file_path = None
        self.source_folder_path = None
        self.sheet_name = None
        self.column_old_name = None
        self.column_new_name = None
        self.status_text = StringVar()
        self.status_text.set("Chưa bắt đầu đổi tên...")

        # Tạo giao diện người dùng
        self.create_widgets()

    def create_widgets(self):
        """Tạo giao diện người dùng"""
        # Khung chọn file Excel
        excel_frame = Frame(self.master, bg="#f2f2f2")
        excel_frame.pack(pady=10)
        Label(excel_frame, text="Chọn file Excel:", fg="#333", bg="#f2f2f2", font=("Arial", 12, "bold")).grid(row=0, column=0, padx=5)
        self.excel_file_label = Label(excel_frame, text="Chưa chọn file", fg="#555", bg="#f2f2f2", font=("Arial", 10))
        self.excel_file_label.grid(row=0, column=1, padx=5)
        self.button_select_excel = Button(excel_frame, text="Chọn file", command=self.select_excel_file, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        self.button_select_excel.grid(row=0, column=2, padx=5)

        # Khung chọn sheet và cột
        sheet_frame = Frame(self.master, bg="#f2f2f2")
        sheet_frame.pack(pady=10)
        Label(sheet_frame, text="Chọn sheet:", fg="#333", bg="#f2f2f2", font=("Arial", 12, "bold")).grid(row=0, column=0, padx=5)
        self.sheet_dropdown = ttk.Combobox(sheet_frame, state="readonly", font=("Arial", 10))
        self.sheet_dropdown.grid(row=0, column=1, padx=5)
        self.sheet_dropdown.bind("<<ComboboxSelected>>", self.select_sheet)

        Label(sheet_frame, text="Cột tên file cũ:", fg="#333", bg="#f2f2f2", font=("Arial", 12, "bold")).grid(row=1, column=0, padx=5)
        self.old_name_dropdown = ttk.Combobox(sheet_frame, state="readonly", font=("Arial", 10))
        self.old_name_dropdown.grid(row=1, column=1, padx=5)
        self.old_name_dropdown.bind("<<ComboboxSelected>>", self.select_old_column)

        Label(sheet_frame, text="Cột tên file mới:", fg="#333", bg="#f2f2f2", font=("Arial", 12, "bold")).grid(row=2, column=0, padx=5)
        self.new_name_dropdown = ttk.Combobox(sheet_frame, state="readonly", font=("Arial", 10))
        self.new_name_dropdown.grid(row=2, column=1, padx=5)
        self.new_name_dropdown.bind("<<ComboboxSelected>>", self.select_new_column)

        # Khung chọn thư mục nguồn
        source_frame = Frame(self.master, bg="#f2f2f2")
        source_frame.pack(pady=10)
        Label(source_frame, text="Chọn thư mục nguồn:", fg="#333", bg="#f2f2f2", font=("Arial", 12, "bold")).grid(row=0, column=0, padx=5)
        self.source_folder_label = Label(source_frame, text="Chưa chọn thư mục", fg="#555", bg="#f2f2f2", font=("Arial", 10))
        self.source_folder_label.grid(row=0, column=1, padx=5)
        self.button_select_source = Button(source_frame, text="Chọn thư mục", command=self.select_source_folder, bg="#2196F3", fg="white", font=("Arial", 10, "bold"))
        self.button_select_source.grid(row=0, column=2, padx=5)

        # Khung nút đổi tên
        rename_frame = Frame(self.master, bg="#42f5ec")
        rename_frame.pack(pady=20)
        self.button_rename = Button(rename_frame, text="Bắt đầu đổi tên", command=self.start_rename_files, state="disabled", bg="#FF5722", fg="white", font=("Arial", 12, "bold"))
        self.button_rename.pack(padx=10)

        # Thanh tiến trình
        self.progress_bar = ttk.Progressbar(rename_frame, orient="horizontal", length=500, mode="determinate")
        self.progress_bar.pack(pady=10)

        # Khung hiển thị trạng thái
        status_frame = Frame(self.master, bg="#f2f2f2")
        status_frame.pack(pady=10, fill="x")
        self.status_label = Label(status_frame, textvariable=self.status_text, fg="#333", bg="#f2f2f2", font=("Arial", 12, "italic"))
        self.status_label.pack(side="left")

        # Khung log
        log_frame = Frame(self.master, bg="#f2f2f2")
        log_frame.pack(pady=10, fill="both", expand=True)
        Label(log_frame, text="Lịch sử đổi tên file:", fg="#333", bg="#f2f2f2", font=("Arial", 12, "bold")).pack(anchor="w")
        self.log_box = Text(log_frame, width=80, height=15, state="disabled", bg="#ffffff", fg="#000", font=("Arial", 10))
        self.log_box.pack(side="left", fill="both", expand=True)

        scrollbar_log = Scrollbar(log_frame)
        scrollbar_log.pack(side="right", fill="y")
        self.log_box.config(yscrollcommand=scrollbar_log.set)
        scrollbar_log.config(command=self.log_box.yview)

    def select_excel_file(self):
        """Chọn file Excel và hiển thị thông tin lên giao diện"""
        self.excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.excel_file_path:
            self.excel_file_label.config(text=os.path.basename(self.excel_file_path))
            self.update_log(f"Đã chọn file Excel: {self.excel_file_path}")
            self.load_excel_data()
        else:
            self.excel_file_label.config(text="Chưa chọn file")

    def load_excel_data(self):
        """Tải sheet và cột từ file Excel"""
        try:
            excel_data = pd.ExcelFile(self.excel_file_path)
            self.sheet_dropdown["values"] = excel_data.sheet_names
            self.sheet_dropdown.set("Chọn sheet")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể đọc file Excel: {e}")

    def select_sheet(self, event):
        """Xử lý khi người dùng chọn sheet"""
        self.sheet_name = self.sheet_dropdown.get()
        self.load_columns()
        self.check_ready_to_rename()

    def load_columns(self):
        """Tải danh sách cột từ sheet được chọn"""
        try:
            data = pd.read_excel(self.excel_file_path, sheet_name=self.sheet_name)
            columns = list(data.columns)
            self.old_name_dropdown["values"] = columns
            self.new_name_dropdown["values"] = columns
            self.old_name_dropdown.set("Chọn cột tên file cũ")
            self.new_name_dropdown.set("Chọn cột tên file mới")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể tải dữ liệu sheet: {e}")

    def select_old_column(self, event):
        """Xử lý khi người dùng chọn cột tên file cũ"""
        self.column_old_name = self.old_name_dropdown.get()
        self.check_ready_to_rename()

    def select_new_column(self, event):
        """Xử lý khi người dùng chọn cột tên file mới"""
        self.column_new_name = self.new_name_dropdown.get()
        self.check_ready_to_rename()

    def select_source_folder(self):
        """Chọn thư mục nguồn và hiển thị thông tin lên giao diện"""
        self.source_folder_path = filedialog.askdirectory()
        if self.source_folder_path:
            self.source_folder_label.config(text=self.source_folder_path)
            self.update_log(f"Đã chọn thư mục nguồn: {self.source_folder_path}")
        else:
            self.source_folder_label.config(text="Chưa chọn thư mục")
        self.check_ready_to_rename()

    def check_ready_to_rename(self):
        """Kích hoạt nút đổi tên nếu đã chọn đầy đủ"""
        if self.excel_file_path and self.source_folder_path and self.sheet_name and self.column_old_name and self.column_new_name:
            self.button_rename.config(state="normal")
        else:
            self.button_rename.config(state="disabled")

    def start_rename_files(self):
        """Bắt đầu đổi tên file từ Excel"""
        self.status_text.set("Đang đổi tên...")
        self.button_rename.config(state="disabled")
        self.rename_files_from_excel()

    def rename_files_from_excel(self):
        """Đổi tên file dựa trên dữ liệu từ Excel"""
        try:
            data = pd.read_excel(self.excel_file_path, sheet_name=self.sheet_name)

            # Kiểm tra cột bắt buộc
            if self.column_old_name not in data.columns or self.column_new_name not in data.columns:
                self.update_log("Lỗi: Sheet không chứa đủ hai cột được chọn.")
                return

            # Lọc dữ liệu hợp lệ
            data = data.dropna(subset=[self.column_old_name, self.column_new_name])

            success_count = 0
            fail_count = 0
            not_found_count = 0

            total_files = len(data)
            self.progress_bar["maximum"] = total_files

            for index, row in data.iterrows():
                old_name = str(row[self.column_old_name]).strip()
                new_name = str(row[self.column_new_name]).strip()

                if not old_name or not new_name:
                    self.update_log(f"[BỎ QUA] Dòng {index + 2}: Tên file cũ hoặc mới trống.")
                    fail_count += 1
                    self.progress_bar["value"] += 1
                    self.master.update()
                    continue

                old_file_path = os.path.join(self.source_folder_path, old_name)
                new_file_path = os.path.join(self.source_folder_path, new_name)

                if not os.path.isfile(old_file_path):
                    self.update_log(f"[KHÔNG TÌM THẤY] Dòng {index + 2}: '{old_name}' không tồn tại.")
                    not_found_count += 1
                    self.progress_bar["value"] += 1
                    self.master.update()
                    continue

                try:
                    os.rename(old_file_path, new_file_path)
                    self.update_log(f"[THÀNH CÔNG] Dòng {index + 2}: '{old_name}' -> '{new_name}'")
                    success_count += 1
                except Exception as e:
                    self.update_log(f"[THẤT BẠI] Dòng {index + 2}: '{old_name}' -> '{new_name}'. Lỗi: {e}")
                    fail_count += 1
                finally:
                    self.progress_bar["value"] += 1
                    self.master.update()

            # Tóm tắt kết quả
            self.update_log("\n===========================================")
            self.update_log("TÓM TẮT KẾT QUẢ:")
            self.update_log(f"Số file đổi tên thành công: {success_count}")
            self.update_log(f"Số file không tìm thấy: {not_found_count}")
            self.update_log(f"Số file bị lỗi: {fail_count}")

            messagebox.showinfo("Hoàn thành", f"Đổi tên hoàn tất!\n\nThành công: {success_count}\nKhông tìm thấy: {not_found_count}\nThất bại: {fail_count}")

        except Exception as e:
            messagebox.showerror("Lỗi", f"Đã xảy ra lỗi: {e}")
        finally:
            self.status_text.set("Quá trình đổi tên đã hoàn thành.")
            self.button_rename.config(state="normal")

    def update_log(self, message):
        """Cập nhật log vào textbox"""
        self.log_box.config(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.yview("end")
        self.log_box.config(state="disabled")

# Khởi tạo ứng dụng Tkinter
if __name__ == "__main__":
    root = Tk()
    app = FileRenameApp(root)
    root.mainloop()
