# Đường dẫn: excel_toolkit/processes/set_label.py
# Phiên bản 3.0 - Cập nhật để hoạt động với ExcelController đã mở sẵn
# Ngày cập nhật: 2025-09-15

import logging
import os
from excel_controller import ExcelController

def run(controller, file_path):
    """
    Quy trình chính: Thêm một nhãn tùy chỉnh vào tất cả các sheet đang
    hiển thị nếu nhãn đó chưa tồn tại.
    """
    shape_name = 'Alliance_Labeling'
    label_text = 'Nissan Confidential C'
    
    logging.info(f"Bắt đầu quy trình dán nhãn cho file: {os.path.basename(file_path)}")
    try:
        # Lấy danh sách sheet hiển thị trực tiếp từ controller
        visible_sheets, _ = controller.get_sheets_visibility()
        
        for sheet_name in visible_sheets:
            logging.debug(f"Đang xử lý sheet: '{sheet_name}'")
            
            # Kiểm tra sự tồn tại của shape trực tiếp từ controller
            if not controller.is_shape_exist(sheet_name, shape_name):
                logging.info(f"Label '{shape_name}' chưa tồn tại, tiến hành tạo mới.")
                # Gọi đến phương thức add_textbox mạnh mẽ của controller
                format_props = {
                    'name': shape_name,
                    'font_name': "Verdana",
                    'font_size': 10,
                    'auto_size': True,
                    'word_wrap': False,
                    'line_visible': True,
                    'line_weight': 1,
                    'line_color': (0, 0, 0)
                }
                controller.add_textbox(
                    sheet_name=sheet_name, text=label_text,
                    top=1, left=1, width=150, height=20,
                    format_properties=format_props
                )
            else:
                logging.info(f"Label '{shape_name}' đã tồn tại. Bỏ qua.")

        logging.info(f"Hoàn tất quy trình dán nhãn cho file: {os.path.basename(file_path)}")

    except Exception as e:
        logging.error(f"Lỗi nghiêm trọng trong quy trình dán nhãn cho file '{file_path}': {e}", exc_info=True)
        raise
