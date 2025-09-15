# Đường dẫn: excel_toolkit/processes/compress_all_images.py
# Phiên bản 2.1 - Cập nhật để nhận tham số chất lượng
# Ngày cập nhật: 2025-09-15

import logging
import os
from excel_controller import ExcelController

def run(controller, file_path, engine='xlwings', quality=220):
    """
    Quy trình nén tất cả hình ảnh trong workbook, cho phép chọn engine và thông số chất lượng.
    """
    logging.info(f"Bắt đầu nén tất cả hình ảnh cho file: {os.path.basename(file_path)} với engine '{engine}' và chất lượng '{quality}'")
    try:
        controller.compress_all_images(file_path, engine=engine, quality=quality)
        logging.info(f"Hoàn tất nén tất cả hình ảnh cho file: {os.path.basename(file_path)}")
    except Exception as e:
        logging.error(f"Lỗi khi nén hình ảnh cho file '{file_path}': {e}", exc_info=True)
        raise
