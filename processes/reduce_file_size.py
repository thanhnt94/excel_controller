# Đường dẫn: excel_toolkit/processes/reduce_file_size.py
# Phiên bản 1.0 - Quy trình tối ưu hóa và giảm dung lượng file
# Ngày cập nhật: 2025-09-12

import logging
import os
from excel_controller import ExcelController

def reduce_file_size(file_path):
    """
    Thực hiện một chuỗi các tác vụ để tối ưu và giảm dung lượng file Excel.

    Bao gồm các bước:
    1. Dọn dẹp định dạng ô thừa.
    2. Nén tất cả hình ảnh.
    3. Làm mới và dọn dẹp cache của Pivot Table.
    """
    logging.info(f"Bắt đầu quy trình giảm dung lượng file cho: {os.path.basename(file_path)}")
    try:
        with ExcelController(visible=False, optimize_performance=True) as controller:
            if not controller.open_workbook(file_path):
                logging.error(f"Không thể mở file, bỏ qua: {os.path.basename(file_path)}")
                return

            logging.info("  -> Bước 1/3: Dọn dẹp định dạng ô thừa...")
            controller.clear_excess_cell_formatting()

            logging.info("  -> Bước 2/3: Nén tất cả hình ảnh...")
            controller.compress_all_images()

            logging.info("  -> Bước 3/3: Dọn dẹp Pivot Table caches...")
            controller.refresh_and_clean_pivot_caches()
            
            controller.save_workbook()
            logging.info(f"Hoàn tất quy trình giảm dung lượng file cho: {os.path.basename(file_path)}")

    except Exception as e:
        logging.error(f"Lỗi nghiêm trọng trong quy trình giảm dung lượng file '{file_path}': {e}", exc_info=True)
        raise

