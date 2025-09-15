# Đường dẫn: excel_toolkit/processes/delete_external_links.py
# Phiên bản 3.0 - Cập nhật để hoạt động với ExcelController đã mở sẵn
# Ngày cập nhật: 2025-09-15

import logging
import os
from excel_controller import ExcelController

def run(controller, file_path):
    """
    Xóa các liên kết ngoài trong workbook.
    """
    logging.info(f"Bắt đầu xóa liên kết ngoài cho file: {os.path.basename(file_path)}")
    try:
        controller.delete_external_links()

        logging.info(f"Hoàn tất xóa liên kết ngoài cho file: {os.path.basename(file_path)}")
    except Exception as e:
        logging.error(f"Lỗi khi xóa liên kết ngoài cho file '{file_path}': {e}", exc_info=True)
        raise
