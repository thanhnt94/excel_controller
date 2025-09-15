# Đường dẫn: excel_toolkit/processes/set_print_settings.py
# Phiên bản 3.0 - Cập nhật để hoạt động với ExcelController đã mở sẵn
# Ngày cập nhật: 2025-09-15

import logging
import os
from excel_controller import ExcelController

def run(controller, file_path):
    """
    Thiết lập trang in thông minh (A3, Ngang) cho tất cả các sheet hiển thị.
    """
    logging.info(f"Bắt đầu thiết lập trang in cho file: {os.path.basename(file_path)}")
    try:
        controller.set_smart_print_settings()
        
        logging.info(f"Hoàn tất thiết lập trang in cho file: {os.path.basename(file_path)}")
    except Exception as e:
        logging.error(f"Lỗi khi thiết lập trang in cho file '{file_path}': {e}", exc_info=True)
        raise
