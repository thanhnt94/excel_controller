# Đường dẫn: excel_toolkit/processes/clear_excess_cell_formatting.py
# Phiên bản 1.0 - Quy trình dọn dẹp định dạng ô thừa
# Ngày cập nhật: 2025-09-15

import logging
import os
from excel_controller import ExcelController

def run(controller, file_path):
    """
    Quy trình dọn dẹp định dạng ô thừa.
    """
    logging.info(f"Bắt đầu dọn dẹp định dạng ô thừa cho file: {os.path.basename(file_path)}")
    try:
        controller.clear_excess_cell_formatting()
        logging.info(f"Hoàn tất dọn dẹp định dạng ô thừa cho file: {os.path.basename(file_path)}")
    except Exception as e:
        logging.error(f"Lỗi khi dọn dẹp định dạng ô thừa cho file '{file_path}': {e}", exc_info=True)
        raise
