# Đường dẫn: excel_toolkit/processes/delete_defined_names.py
# Phiên bản 3.0 - Cập nhật để hoạt động với ExcelController đã mở sẵn
# Ngày cập nhật: 2025-09-15

import logging
import os
from excel_controller import ExcelController

def run(controller, file_path):
    """
    Xóa Defined Name trong workbook (an toàn, không xóa thiết lập in).
    """
    logging.info(f"Bắt đầu xóa Defined Name cho file: {os.path.basename(file_path)}")
    try:
        controller.delete_defined_names()

        logging.info(f"Hoàn tất xóa Defined Name cho file: {os.path.basename(file_path)}")
    except Exception as e:
        logging.error(f"Lỗi khi xóa Defined Name cho file '{file_path}': {e}", exc_info=True)
        raise
