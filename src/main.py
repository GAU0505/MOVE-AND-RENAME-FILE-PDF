import os
import subprocess
from tkinter import Tk, Button, Label, Frame

class MainApp:
    def __init__(self, master):
        self.master = master
        self.master.title("QUẢN LÝ FILE PDF - HỒ XUÂN ÁNH")
        self.master.geometry("500x300")
        self.master.resizable(False, False)
        self.master.config(bg="#42f5ec")

        # Giao diện chính
        self.create_widgets()

    def create_widgets(self):
        """Tạo giao diện chính"""
        # Tiêu đề
        title_label = Label(
            self.master,
            text="CHỌN CHỨC NĂNG",
            font=("Arial", 16, "bold"),
            fg="#333",
            bg="#42f5ec"
        )
        title_label.pack(pady=20)

        # Khung chứa nút
        button_frame = Frame(self.master, bg="#42f5ec")
        button_frame.pack(pady=20)

        # Nút RENAME
        rename_button = Button(
            button_frame,
            text="ĐỔI TÊN FILE PDF (RENAME)",
            font=("Arial", 12, "bold"),
            bg="#4CAF50",
            fg="white",
            width=30,
            height=2,
            command=self.run_rename
        )
        rename_button.pack(pady=10)

        # Nút MOVE
        move_button = Button(
            button_frame,
            text="DI CHUYỂN FILE PDF (MOVE)",
            font=("Arial", 12, "bold"),
            bg="#2196F3",
            fg="white",
            width=30,
            height=2,
            command=self.run_move
        )
        move_button.pack(pady=10)

        # Ghi chú
        note_label = Label(
            self.master,
            text="Vui lòng chọn chức năng để tiếp tục",
            font=("Arial", 10, "italic"),
            fg="#555",
            bg="#42f5ec"
        )
        note_label.pack(pady=10)

    def run_rename(self):
        """Khởi chạy chức năng Rename"""
        self.run_external_script("RENAME.py")

    def run_move(self):
        """Khởi chạy chức năng Move"""
        self.run_external_script("MOVE.py")

    def run_external_script(self, script_name):
        """Thực thi file script bên ngoài"""
        if os.path.exists(script_name):
            subprocess.Popen(["python", script_name])
        else:
            print(f"Lỗi: Không tìm thấy file {script_name}")

# Khởi tạo giao diện chính
if __name__ == "__main__":
    root = Tk()
    app = MainApp(root)
    root.mainloop()
