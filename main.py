# Đường dẫn: excel_toolkit/main.py
# Phiên bản 31.0 - Tái cấu trúc, tách UI và logic
# Ngày cập nhật: 2025-09-16

import customtkinter
import logging
from datetime import datetime
import os
from app_controller import AppController

# --- Hệ thống log ---
LOG_DIR = "logs"
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_FILENAME = os.path.join(LOG_DIR, f"log_{TIMESTAMP}.log")

def configure_logging(level=logging.INFO):
    """
    Khởi tạo hệ thống ghi log một lần duy nhất khi chương trình bắt đầu.
    """
    log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    
    root_logger = logging.getLogger()
    root_logger.setLevel(level)

    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)

    file_handler = logging.FileHandler(LOG_FILENAME, 'w', 'utf-8')
    file_handler.setFormatter(log_formatter)
    root_logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(log_formatter)
    root_logger.addHandler(stream_handler)
    
    logging.info(f"Hệ thống ghi log đã được khởi tạo. Mức độ: {logging.getLevelName(level)}. File: {LOG_FILENAME}")

if __name__ == "__main__":
    configure_logging(logging.INFO)
    
    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("blue")

    root = customtkinter.CTk()
    controller = AppController(root)
    root.mainloop()

