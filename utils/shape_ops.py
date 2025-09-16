# Đường dẫn: excel_toolkit/utils/shape_ops.py
# Phiên bản 6.0 - Loại bỏ các hàm nén ảnh cũ, không hiệu quả
# Ngày cập nhật: 2025-09-15

import logging
import xlwings as xw
import os

# ======================================================================
# --- Nhóm 1: Lấy thông tin & Trạng thái ---
# ======================================================================

def is_shape_exist(wb, sheet_name, shape_name):
    """
    Kiểm tra xem một shape có tồn tại trong sheet hay không.
    """
    logging.debug(f"Bắt đầu kiểm tra sự tồn tại của shape '{shape_name}' trong sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        for shape in sheet.api.Shapes:
            if shape.Name == shape_name:
                logging.debug(f"  -> Đã tìm thấy shape '{shape_name}'.")
                return True
        logging.debug(f"  -> Không tìm thấy shape '{shape_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi kiểm tra sự tồn tại của shape '{shape_name}': {e}")
        return False

def get_all_shape_names(wb, sheet_name):
    """
    Trả về một danh sách chứa tên của tất cả các shape có trên một sheet.
    """
    logging.debug(f"Bắt đầu lấy danh sách tên các shape từ sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        shape_names = [shape.name for shape in sheet.shapes]
        logging.info(f"Đã tìm thấy {len(shape_names)} shape trong sheet '{sheet_name}'.")
        return shape_names
    except Exception as e:
        logging.error(f"Lỗi khi lấy danh sách shape: {e}")
        return []

# ======================================================================
# --- Nhóm 2: Thêm & Sửa đổi Đối tượng ---
# ======================================================================

def add_textbox(wb, sheet_name, text, top, left, width, height, format_properties=None):
    """
    Tạo một textbox mới với các thuộc tính định dạng tùy chỉnh.
    """
    logging.debug(f"Bắt đầu thêm textbox vào sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        shape = sheet.shapes.add_textbox(text, top, left, width, height)
        if format_properties:
            logging.debug("  -> Áp dụng các thuộc tính định dạng tùy chỉnh...")
            if 'name' in format_properties: shape.name = format_properties['name']
            if 'font_name' in format_properties: shape.text_frame.font.name = format_properties['font_name']
            if 'font_size' in format_properties: shape.text_frame.font.size = format_properties['font_size']
            if 'bold' in format_properties: shape.text_frame.font.bold = format_properties['bold']
            if 'italic' in format_properties: shape.text_frame.font.italic = format_properties['italic']
            if 'text_color' in format_properties: shape.text_frame.font.color = format_properties['text_color']
            if 'auto_size' in format_properties: shape.text_frame.auto_size = format_properties['auto_size']
        logging.info(f"Đã thêm textbox '{shape.name}' thành công.")
        return shape.name
    except Exception as e:
        logging.error(f"Lỗi khi thêm textbox: {e}")
        return None

def add_picture(wb, sheet_name, image_path, top, left, width=None, height=None, name=None):
    """
    Chèn một hình ảnh từ một đường dẫn file vào sheet.
    """
    logging.debug(f"Bắt đầu chèn ảnh '{os.path.basename(image_path)}' vào sheet '{sheet_name}'.")
    if not os.path.exists(image_path):
        logging.error(f"  -> Lỗi: Không tìm thấy file ảnh tại đường dẫn: {image_path}")
        return None
    try:
        sheet = wb.sheets[sheet_name]
        picture = sheet.pictures.add(image_path, top=top, left=left, width=width, height=height, name=name)
        logging.info(f"Đã chèn ảnh '{picture.name}' thành công.")
        return picture.name
    except Exception as e:
        logging.error(f"Lỗi khi chèn ảnh: {e}")
        return None

def edit_textbox(wb, sheet_name, shape_name, new_text):
    """
    Chỉnh sửa nội dung văn bản của một textbox.
    """
    logging.debug(f"Bắt đầu chỉnh sửa textbox '{shape_name}' trong sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        shape = sheet.shapes[shape_name]
        shape.text = new_text
        logging.info(f"Đã cập nhật thành công nội dung của shape '{shape_name}'.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi chỉnh sửa textbox '{shape_name}': {e}")
        return False

def copy_shape(source_wb, source_sheet_name, target_wb, target_sheet_name, shape_name):
    """
    Sao chép một shape từ một sheet nguồn sang một sheet đích, có thể giữa các workbook khác nhau.
    """
    logging.debug(f"Bắt đầu sao chép shape '{shape_name}' từ '{source_wb.name}/{source_sheet_name}' sang '{target_wb.name}/{target_sheet_name}'.")
    try:
        source_sheet = source_wb.sheets[source_sheet_name]
        target_sheet = target_wb.sheets[target_sheet_name]
        shape_to_copy = source_sheet.shapes[shape_name]
        shape_to_copy.api.Copy()
        target_sheet.api.Paste()
        pasted_shape = target_sheet.shapes[-1]
        logging.info(f"Đã dán shape '{pasted_shape.name}' thành công vào sheet '{target_sheet_name}'.")
        return pasted_shape.name
    except Exception as e:
        logging.error(f"Lỗi khi sao chép shape '{shape_name}': {e}")
        return None

# ======================================================================
# --- Nhóm 3: Vị trí & Kích thước ---
# ======================================================================

def move_shape(wb, sheet_name, shape_name, top, left):
    """
    Di chuyển một shape đến một vị trí mới.
    """
    logging.debug(f"Bắt đầu di chuyển shape '{shape_name}' đến vị trí ({top}, {left}).")
    try:
        sheet = wb.sheets[sheet_name]
        shape = sheet.shapes[shape_name]
        shape.top = top
        shape.left = left
        logging.info(f"Đã di chuyển thành công shape '{shape_name}'.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi di chuyển shape '{shape_name}': {e}")
        return False

def resize_shape(wb, sheet_name, shape_name, width, height):
    """
    Thay đổi kích thước của một shape.
    """
    logging.debug(f"Bắt đầu thay đổi kích thước shape '{shape_name}' thành ({width}x{height}).")
    try:
        sheet = wb.sheets[sheet_name]
        shape = sheet.shapes[shape_name]
        shape.width = width
        shape.height = height
        logging.info(f"Đã thay đổi kích thước thành công cho shape '{shape_name}'.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi thay đổi kích thước shape '{shape_name}': {e}")
        return False

# ======================================================================
# --- Nhóm 4: Xóa ---
# ======================================================================

def delete_shape(wb, sheet_name, shape_name):
    """
    Xóa một shape cụ thể khỏi sheet.
    """
    logging.debug(f"Bắt đầu xóa shape '{shape_name}' khỏi sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        sheet.shapes[shape_name].delete()
        logging.info(f"Đã xóa thành công shape '{shape_name}'.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi xóa shape '{shape_name}': {e}")
        return False
