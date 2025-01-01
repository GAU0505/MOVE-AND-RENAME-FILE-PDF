import os
import shutil
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
from tqdm import tqdm
import logging
from tkinter import Tk, filedialog, messagebox, Label, Button, StringVar, Frame, Text, Scrollbar, VERTICAL, RIGHT, Y
import json
import threading

# Bật log chi tiết để debug
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

HISTORY_FILE = "folder_history.json"  # Tệp tin lưu lịch sử

class FileMoveHandler(FileSystemEventHandler):
    def __init__(self, source_folder, destination_folder, status_var, history_text):
        self.source_folder = source_folder
        self.destination_folder = destination_folder
        self.moved_files = set()
        self.status_var = status_var
        self.history_text = history_text

    def on_created(self, event):
        logging.info(f"Đã phát hiện sự kiện: {event.event_type}, Đường dẫn: {event.src_path}")
        if not event.is_directory and event.src_path.lower().endswith('.pdf'):
            filename = os.path.basename(event.src_path)
            source_file = event.src_path
            target_file = os.path.join(self.destination_folder, filename)

            if filename in self.moved_files:
                logging.warning(f"Bỏ qua file đã di chuyển: {filename}")
                return

            logging.info(f"Đang xử lý file: {event.src_path}")
            while not self.is_file_ready(source_file):
                logging.debug(f"Đang chờ file sẵn sàng: {filename}")
                time.sleep(1)

            if not self.is_file_valid(source_file):
                logging.warning(f"File không hợp lệ hoặc bị lỗi: {filename}")
                self.status_var.set(f"File không hợp lệ hoặc bị lỗi: {filename}")
                self.history_text.insert('end', f"File không hợp lệ hoặc bị lỗi: {filename}\n")
                return

            try:
                logging.info(f"Đang di chuyển {filename}...")
                self.status_var.set(f"Đang di chuyển {filename}...")
                shutil.copy2(source_file, target_file)
                os.remove(source_file)
                logging.info(f"Đã di chuyển: {source_file} -> {target_file}")
                self.moved_files.add(filename)
                self.status_var.set(f"Đã di chuyển: {filename}")
                self.history_text.insert('end', f"Đã di chuyển: {filename}\n")
            except Exception as e:
                logging.error(f"Không thể di chuyển {source_file}: {e}")
                self.status_var.set(f"Lỗi: {e}")
                self.history_text.insert('end', f"Lỗi: {e}\n")
        else:
            logging.info(f"Bỏ qua file không phải PDF: {event.src_path}")

    @staticmethod
    def is_file_ready(file_path):
        try:
            with open(file_path, 'rb'):
                return True
        except IOError:
            return False

    @staticmethod
    def is_file_valid(file_path):
        """Kiểm tra xem file có dung lượng và không bị lỗi."""
        try:
            if os.path.getsize(file_path) == 0:
                return False
            with open(file_path, 'rb') as f:
                f.read()
            return True
        except Exception as e:
            logging.error(f"Lỗi khi kiểm tra file {file_path}: {e}")
            return False


def process_existing_files(source_folder, destination_folder, status_var, history_text):
    logging.info(f"Đang kiểm tra các file hiện có trong {source_folder}")
    for file_name in os.listdir(source_folder):
        source_file = os.path.join(source_folder, file_name)
        if os.path.isfile(source_file) and file_name.lower().endswith('.pdf'):
            if not FileMoveHandler.is_file_valid(source_file):
                logging.warning(f"File không hợp lệ hoặc bị lỗi: {file_name}")
                status_var.set(f"File không hợp lệ hoặc bị lỗi: {file_name}")
                history_text.insert('end', f"File không hợp lệ hoặc bị lỗi: {file_name}\n")
                continue
            target_file = os.path.join(destination_folder, file_name)
            try:
                if not os.path.exists(target_file):
                    logging.info(f"Đang di chuyển file tồn tại: {file_name}")
                    status_var.set(f"Đang di chuyển file tồn tại: {file_name}")
                    shutil.move(source_file, target_file)
                    history_text.insert('end', f"Đã di chuyển: {file_name}\n")
                else:
                    logging.info(f"File đã tồn tại tại đích, bỏ qua: {file_name}")
            except Exception as e:
                logging.error(f"Không thể di chuyển file {file_name}: {e}")
                status_var.set(f"Lỗi: {e}")
                history_text.insert('end', f"Lỗi: {e}\n")


def watch_and_move(source_folder, destination_folder, status_var, history_text):
    if not os.path.exists(source_folder):
        logging.error(f"Lỗi: Thư mục nguồn {source_folder} không tồn tại!")
        status_var.set(f"Lỗi: Thư mục nguồn {source_folder} không tồn tại!")
        return

    if not os.path.exists(destination_folder):
        logging.info(f"Thư mục đích {destination_folder} không tồn tại, đang tạo mới...")
        os.makedirs(destination_folder)

    process_existing_files(source_folder, destination_folder, status_var, history_text)

    logging.info(f"Đang giám sát thư mục: {source_folder}")
    logging.info(f"Tất cả file PDF sẽ được di chuyển vào: {destination_folder}")
    status_var.set(f"Đang giám sát thư mục: {source_folder}")

    event_handler = FileMoveHandler(source_folder, destination_folder, status_var, history_text)
    observer = Observer()
    observer.schedule(event_handler, source_folder, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        logging.info("\nDừng giám sát.")
        status_var.set("Dừng giám sát.")
    observer.join()

    # Thêm đoạn mã này để tự động thoát sau khi di chuyển các tệp hiện có
    logging.info("Đã di chuyển tất cả các tệp hiện có. Thoát chương trình.")
    status_var.set("Đã di chuyển tất cả các tệp hiện có. Thoát chương trình.")
    root.quit()


def save_history(source_folder, destination_folder):
    """Lưu lịch sử thư mục vào tệp JSON."""
    history = {"source_folder": source_folder, "destination_folder": destination_folder}
    with open(HISTORY_FILE, 'w') as f:
        json.dump(history, f)
    logging.info("Đã lưu lịch sử thư mục.")


def load_history():
    """Đọc lịch sử thư mục từ tệp JSON."""
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, 'r') as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                logging.error("Lỗi đọc tệp lịch sử. Sử dụng giá trị mặc định.")
    return None


def select_folders():
    """Cho phép người dùng chọn Folder mới."""
    root = Tk()
    root.withdraw()  # Ẩn cửa sổ chính của Tkinter
    logging.info("Hãy chọn Folder nguồn...")
    source_folder = filedialog.askdirectory(title="Chọn Folder nguồn")
    if not source_folder:
        logging.error("Bạn chưa chọn Folder nguồn. Thoát chương trình.")
        return None, None

    logging.info("Hãy chọn Folder Save...")
    destination_folder = filedialog.askdirectory(title="Chọn Folder Save")
    if not destination_folder:
        logging.error("Bạn chưa chọn Folder Save. Thoát chương trình.")
        return None, None

    save_history(source_folder, destination_folder)
    return source_folder, destination_folder


def start_watching(source_folder, destination_folder, status_var, history_text, source_label, destination_label):
    source_label.config(text=f"Lịch sử nguồn: {source_folder}")
    destination_label.config(text=f"Lịch sử đích: {destination_folder}")
    threading.Thread(target=watch_and_move, args=(source_folder, destination_folder, status_var, history_text)).start()


def main():
    root = Tk()
    root.title("PDF File Mover")
    root.geometry("700x500")
    root.configure(bg="#f0f0f0")

    status_var = StringVar()
    status_var.set("Chọn thư mục để bắt đầu.")

    frame = Frame(root, bg="#f0f0f0")
    frame.pack(pady=20)

    Label(frame, text="PDF File Mover", font=("Helvetica", 20, "bold"), bg="#f0f0f0", fg="#333").pack(pady=10)
    Label(frame, textvariable=status_var, bg="#f0f0f0", fg="#333").pack(pady=10)

    history_text = Text(frame, height=10, width=70, wrap='word', bg="#fff", fg="#333", font=("Helvetica", 10))
    history_text.pack(pady=10)
    scrollbar = Scrollbar(frame, orient=VERTICAL, command=history_text.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    history_text.config(yscrollcommand=scrollbar.set)

    history = load_history()
    source_label = Label(frame, text="Lịch sử nguồn: Không có lịch sử", bg="#f0f0f0", fg="#333")
    destination_label = Label(frame, text="Lịch sử đích: Không có lịch sử", bg="#f0f0f0", fg="#333")
    source_label.pack(pady=5)
    destination_label.pack(pady=5)
    if history:
        source_folder = history.get("source_folder", "Không có lịch sử")
        destination_folder = history.get("destination_folder", "Không có lịch sử")
        source_label.config(text=f"Lịch sử nguồn: {source_folder}")
        destination_label.config(text=f"Lịch sử đích: {destination_folder}")

    def on_select_new_folders():
        source_folder, destination_folder = select_folders()
        if source_folder and destination_folder:
            status_var.set(f"Đang giám sát thư mục: {source_folder}")
            start_watching(source_folder, destination_folder, status_var, history_text, source_label, destination_label)

    def on_select_old_folders():
        history = load_history()
        if history:
            source_folder = history.get("source_folder", "")
            destination_folder = history.get("destination_folder", "")
            if source_folder and destination_folder:
                status_var.set(f"Đang giám sát thư mục: {source_folder}")
                start_watching(source_folder, destination_folder, status_var, history_text, source_label, destination_label)
            else:
                messagebox.showerror("Lỗi", "Không tìm thấy lịch sử thư mục.")
        else:
            messagebox.showerror("Lỗi", "Không tìm thấy lịch sử thư mục.")

    Button(frame, text="Chọn thư mục mới", command=on_select_new_folders, bg="#4CAF50", fg="white", font=("Helvetica", 12), width=20).pack(pady=5)
    Button(frame, text="Sử dụng thư mục cũ", command=on_select_old_folders, bg="#2196F3", fg="white", font=("Helvetica", 12), width=20).pack(pady=5)

    root.mainloop()


if __name__ == "__main__":
    main()
