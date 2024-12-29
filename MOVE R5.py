import os
import shutil
import pandas as pd
from tkinter import Tk, filedialog, messagebox, StringVar, Text, Scrollbar, Frame, Label, Button, ttk

class FileMoverApp:
    def __init__(self, master):
        self.master = master
        self.master.title("DI CHUYỂN FILE PDF - HỒ XUÂN ÁNH")
        self.master.geometry("900x700")
        self.master.resizable(False, False)
        self.master.config(bg="#42f5ec")

        # Biến trạng thái
        self.excel_file_path = None
        self.source_base_folder = None
        self.sheet_name = None
        self.column_file_name = None
        self.column_destination = None
        self.status_text = StringVar()
        self.status_text.set("Chưa bắt đầu di chuyển...")

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

        Label(sheet_frame, text="Cột tên file:", fg="#333", bg="#f2f2f2", font=("Arial", 12, "bold")).grid(row=1, column=0, padx=5)
        self.file_name_dropdown = ttk.Combobox(sheet_frame, state="readonly", font=("Arial", 10))
        self.file_name_dropdown.grid(row=1, column=1, padx=5)
        self.file_name_dropdown.bind("<<ComboboxSelected>>", self.select_file_name_column)

        Label(sheet_frame, text="Cột đích đến:", fg="#333", bg="#f2f2f2", font=("Arial", 12, "bold")).grid(row=2, column=0, padx=5)
        self.destination_dropdown = ttk.Combobox(sheet_frame, state="readonly", font=("Arial", 10))
        self.destination_dropdown.grid(row=2, column=1, padx=5)
        self.destination_dropdown.bind("<<ComboboxSelected>>", self.select_destination_column)

        # Khung chọn thư mục nguồn
        source_frame = Frame(self.master, bg="#f2f2f2")
        source_frame.pack(pady=10)
        Label(source_frame, text="Chọn thư mục nguồn:", fg="#333", bg="#f2f2f2", font=("Arial", 12, "bold")).grid(row=0, column=0, padx=5)
        self.source_folder_label = Label(source_frame, text="Chưa chọn thư mục", fg="#555", bg="#f2f2f2", font=("Arial", 10))
        self.source_folder_label.grid(row=0, column=1, padx=5)
        self.button_select_source = Button(source_frame, text="Chọn thư mục", command=self.select_source_folder, bg="#2196F3", fg="white", font=("Arial", 10, "bold"))
        self.button_select_source.grid(row=0, column=2, padx=5)

        # Khung nút di chuyển
        move_frame = Frame(self.master, bg="#42f5ec")
        move_frame.pack(pady=20)
        self.button_move = Button(move_frame, text="Bắt đầu di chuyển", command=self.start_move_files, state="disabled", bg="#FF5722", fg="white", font=("Arial", 12, "bold"))
        self.button_move.pack(padx=10)

        # Thanh tiến trình
        self.progress_bar = ttk.Progressbar(move_frame, orient="horizontal", length=500, mode="determinate")
        self.progress_bar.pack(pady=10)

        # Khung hiển thị trạng thái
        status_frame = Frame(self.master, bg="#f2f2f2")
        status_frame.pack(pady=10, fill="x")
        self.status_label = Label(status_frame, textvariable=self.status_text, fg="#333", bg="#f2f2f2", font=("Arial", 12, "italic"))
        self.status_label.pack(side="left")

        # Khung log
        log_frame = Frame(self.master, bg="#f2f2f2")
        log_frame.pack(pady=10, fill="both", expand=True)
        Label(log_frame, text="Lịch sử di chuyển file:", fg="#333", bg="#f2f2f2", font=("Arial", 12, "bold")).pack(anchor="w")
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
        """Tải sheet từ file Excel"""
        try:
            excel_data = pd.ExcelFile(self.excel_file_path)
            self.sheet_dropdown["values"] = excel_data.sheet_names
            self.sheet_dropdown.set("Chọn sheet")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể đọc file Excel: {e}")

    def select_sheet(self, event):
        """Xử lý khi chọn sheet"""
        self.sheet_name = self.sheet_dropdown.get()
        self.load_columns()
        self.check_ready_to_move()

    def load_columns(self):
        """Tải danh sách cột từ sheet"""
        try:
            data = pd.read_excel(self.excel_file_path, sheet_name=self.sheet_name)
            columns = list(data.columns)
            self.file_name_dropdown["values"] = columns
            self.destination_dropdown["values"] = columns
            self.file_name_dropdown.set("Chọn cột tên file")
            self.destination_dropdown.set("Chọn cột đích đến")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể tải dữ liệu sheet: {e}")

    def select_file_name_column(self, event):
        """Xử lý khi chọn cột tên file"""
        self.column_file_name = self.file_name_dropdown.get()
        self.check_ready_to_move()

    def select_destination_column(self, event):
        """Xử lý khi chọn cột đích đến"""
        self.column_destination = self.destination_dropdown.get()
        self.check_ready_to_move()

    def select_source_folder(self):
        """Chọn thư mục nguồn"""
        self.source_base_folder = filedialog.askdirectory()
        if self.source_base_folder:
            self.source_folder_label.config(text=self.source_base_folder)
            self.update_log(f"Đã chọn thư mục nguồn: {self.source_base_folder}")
        else:
            self.source_folder_label.config(text="Chưa chọn thư mục")
        self.check_ready_to_move()

    def check_ready_to_move(self):
        """Kích hoạt nút di chuyển nếu đầy đủ thông tin"""
        if self.excel_file_path and self.source_base_folder and self.sheet_name and self.column_file_name and self.column_destination:
            self.button_move.config(state="normal")
        else:
            self.button_move.config(state="disabled")

    def start_move_files(self):
        """Bắt đầu di chuyển file"""
        self.status_text.set("Đang di chuyển...")
        self.button_move.config(state="disabled")
        self.move_files_from_excel()

    def move_files_from_excel(self):
        """Di chuyển file từ Excel"""
        try:
            data = pd.read_excel(self.excel_file_path, sheet_name=self.sheet_name)
            if self.column_file_name not in data.columns or self.column_destination not in data.columns:
                self.update_log("Lỗi: Sheet không chứa đủ hai cột được chọn.")
                return

            data = data.dropna(subset=[self.column_file_name, self.column_destination])
            total_files = len(data)
            self.progress_bar["maximum"] = total_files

            success_count = 0
            fail_count = 0

            for index, row in data.iterrows():
                file_name = str(row[self.column_file_name]).strip()
                destination_folder = str(row[self.column_destination]).strip()

                source_path = os.path.join(self.source_base_folder, file_name)
                destination_path = os.path.join(destination_folder, file_name)

                try:
                    os.makedirs(destination_folder, exist_ok=True)
                    if os.path.isfile(source_path):
                        shutil.move(source_path, destination_path)
                        self.update_log(f"[THÀNH CÔNG] {file_name} -> {destination_folder}")
                        success_count += 1
                    else:
                        self.update_log(f"[KHÔNG TÌM THẤY] {file_name}")
                        fail_count += 1
                except Exception as e:
                    self.update_log(f"[LỖI] {file_name}: {e}")
                    fail_count += 1
                finally:
                    self.progress_bar["value"] += 1
                    self.master.update()

            self.update_log("\n===========================================")
            self.update_log(f"Thành công: {success_count}, Lỗi: {fail_count}")
            messagebox.showinfo("Hoàn thành", f"Thành công: {success_count}\nLỗi: {fail_count}")

        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi trong quá trình di chuyển: {e}")
        finally:
            self.status_text.set("Hoàn thành.")
            self.button_move.config(state="normal")

    def update_log(self, message):
        """Cập nhật log"""
        self.log_box.config(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.yview("end")
        self.log_box.config(state="disabled")

# Khởi chạy ứng dụng
if __name__ == "__main__":
    root = Tk()
    app = FileMoverApp(root)
    root.mainloop()
