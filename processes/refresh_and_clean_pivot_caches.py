# Đường dẫn: excel_toolkit/processes/refresh_and_clean_pivot_caches.py
# Phiên bản 1.0 - Quy trình dọn dẹp Pivot Table caches
# Ngày cập nhật: 2025-09-15

import logging
import os
from excel_controller import ExcelController

def run(controller, file_path):
    """
    Quy trình làm mới và dọn dẹp Pivot Table caches.
    """
    logging.info(f"Bắt đầu làm mới và dọn dẹp Pivot Table caches cho file: {os.path.basename(file_path)}")
    try:
        controller.refresh_and_clean_pivot_caches()
        logging.info(f"Hoàn tất làm mới và dọn dẹp Pivot Table caches cho file: {os.path.basename(file_path)}")
    except Exception as e:
        logging.error(f"Lỗi khi dọn dẹp Pivot Table caches cho file '{file_path}': {e}", exc_info=True)
        raise
