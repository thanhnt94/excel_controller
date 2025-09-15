# Đường dẫn: excel_toolkit/utils/range_ops.py
# Phiên bản 2.0 - Hoàn thiện các hàm thao tác dữ liệu và định dạng
# Ngày cập nhật: 2025-09-12

import logging
import xlwings as xw

# ======================================================================
# --- Nhóm 1: Đọc & Ghi Dữ liệu ---
# ======================================================================

def get_cell_value(wb, sheet_name, cell_address):
    """Đọc giá trị từ một ô duy nhất."""
    logging.debug(f"Bắt đầu đọc giá trị từ ô '{cell_address}' trên sheet '{sheet_name}'.")
    try:
        value = wb.sheets[sheet_name].range(cell_address).value
        logging.info(f"Giá trị tại ô '{cell_address}' là: {value}")
        return value
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return None
    except Exception as e:
        logging.error(f"Lỗi khi đọc giá trị từ ô '{cell_address}': {e}")
        return None

def set_cell_value(wb, sheet_name, cell_address, value):
    """Ghi một giá trị mới vào một ô duy nhất."""
    logging.debug(f"Bắt đầu ghi giá trị '{value}' vào ô '{cell_address}' trên sheet '{sheet_name}'.")
    try:
        wb.sheets[sheet_name].range(cell_address).value = value
        logging.info(f"Đã ghi giá trị vào ô '{cell_address}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi ghi giá trị vào ô '{cell_address}': {e}")
        return False

def get_range_values(wb, sheet_name, range_address):
    """Đọc dữ liệu từ một vùng và trả về dưới dạng danh sách 2 chiều."""
    logging.debug(f"Bắt đầu đọc giá trị từ vùng '{range_address}' trên sheet '{sheet_name}'.")
    try:
        values = wb.sheets[sheet_name].range(range_address).options(ndim=2).value
        logging.info(f"Đã đọc thành công dữ liệu từ vùng '{range_address}'.")
        return values
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return None
    except Exception as e:
        logging.error(f"Lỗi khi đọc giá trị từ vùng '{range_address}': {e}")
        return None

def set_range_values(wb, sheet_name, start_cell, values):
    """Ghi một danh sách 2 chiều vào sheet, bắt đầu từ một ô."""
    logging.debug(f"Bắt đầu ghi một khối dữ liệu vào sheet '{sheet_name}' tại ô '{start_cell}'.")
    try:
        wb.sheets[sheet_name].range(start_cell).options(expand='table').value = values
        logging.info(f"Đã ghi khối dữ liệu vào sheet '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi ghi khối dữ liệu: {e}")
        return False

def get_last_row(wb, sheet_name, column=1):
    """Tìm hàng cuối cùng có dữ liệu trong một cột cụ thể."""
    logging.debug(f"Bắt đầu tìm hàng cuối cùng có dữ liệu trong cột {column} của sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        last_row = sheet.range(sheet.cells.rows.count, column).end('up').row
        logging.info(f"Hàng cuối cùng có dữ liệu trong cột {column} là: {last_row}")
        return last_row
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return 0
    except Exception as e:
        logging.error(f"Lỗi khi tìm hàng cuối cùng: {e}")
        return 0

# ======================================================================
# --- Nhóm 2: Định dạng & Bố cục ---
# ======================================================================

def format_range(wb, sheet_name, range_address, format_properties):
    """Áp dụng các thuộc tính định dạng cho một vùng."""
    logging.debug(f"Bắt đầu áp dụng định dạng cho vùng '{range_address}' trên sheet '{sheet_name}'.")
    try:
        rng = wb.sheets[sheet_name].range(range_address)
        api = rng.api
        
        if 'bold' in format_properties:
            api.Font.Bold = format_properties['bold']
        if 'italic' in format_properties:
            api.Font.Italic = format_properties['italic']
        if 'underline' in format_properties:
            api.Font.Underline = format_properties['underline']
        if 'color' in format_properties: # (R, G, B)
            api.Font.Color = format_properties['color'][0] + (format_properties['color'][1] * 256) + (format_properties['color'][2] * 256 * 256)
        if 'bg_color' in format_properties: # (R, G, B)
            api.Interior.Color = format_properties['bg_color'][0] + (format_properties['bg_color'][1] * 256) + (format_properties['bg_color'][2] * 256 * 256)
        if 'align_h' in format_properties: # 'left', 'center', 'right'
            align_map = {'left': -4131, 'center': -4108, 'right': -4152}
            api.HorizontalAlignment = align_map.get(format_properties['align_h'].lower(), -4131)
        if 'align_v' in format_properties: # 'top', 'center', 'bottom'
            align_map = {'top': -4160, 'center': -4108, 'bottom': -4107}
            api.VerticalAlignment = align_map.get(format_properties['align_v'].lower(), -4107)

        logging.info(f"Đã áp dụng định dạng cho vùng '{range_address}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi định dạng vùng '{range_address}': {e}")
        return False

def merge_cells(wb, sheet_name, range_address):
    """Gộp các ô trong một vùng lại với nhau."""
    logging.debug(f"Bắt đầu gộp ô cho vùng '{range_address}' trên sheet '{sheet_name}'.")
    try:
        wb.sheets[sheet_name].range(range_address).merge()
        logging.info(f"Đã gộp ô vùng '{range_address}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi gộp ô: {e}")
        return False

def unmerge_cells(wb, sheet_name, range_address):
    """Tách các ô đã được gộp."""
    logging.debug(f"Bắt đầu tách ô cho vùng '{range_address}' trên sheet '{sheet_name}'.")
    try:
        wb.sheets[sheet_name].range(range_address).unmerge()
        logging.info(f"Đã tách ô vùng '{range_address}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi tách ô: {e}")
        return False

def autofit_columns(wb, sheet_name, range_address=None):
    """Tự động điều chỉnh độ rộng cột để vừa với nội dung."""
    logging.debug(f"Bắt đầu autofit cột cho sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        target_range = sheet.range(range_address) if range_address else sheet.used_range
        target_range.columns.autofit()
        logging.info(f"Đã autofit cột cho vùng '{target_range.address}' trên sheet '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi autofit cột: {e}")
        return False

def freeze_panes(wb, sheet_name, cell_address):
    """Cố định các hàng và cột dựa trên một ô."""
    logging.debug(f"Bắt đầu cố định (freeze panes) tại ô '{cell_address}' trên sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        sheet.activate()
        sheet.range(cell_address).select()
        wb.app.api.ActiveWindow.FreezePanes = True
        logging.info(f"Đã cố định màn hình tại ô '{cell_address}' trên sheet '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi cố định màn hình: {e}")
        return False

# ======================================================================
# --- Nhóm 3: Ghi chú, Hyperlink & Sắp xếp ---
# ======================================================================

def add_comment(wb, sheet_name, cell_address, text):
    """Thêm một ghi chú (comment) vào một ô cụ thể."""
    logging.debug(f"Bắt đầu thêm comment tại ô '{cell_address}' trên sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        sheet.range(cell_address).add_comment(text)
        logging.info(f"Đã thêm comment tại ô '{cell_address}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi thêm comment: {e}")
        return False

def add_hyperlink(wb, sheet_name, cell_address, link_address, display_text=None):
    """Tạo một siêu liên kết trong một ô."""
    logging.debug(f"Bắt đầu thêm hyperlink tại ô '{cell_address}' đến '{link_address}'.")
    try:
        sheet = wb.sheets[sheet_name]
        cell = sheet.range(cell_address)
        sheet.api.Hyperlinks.Add(
            Anchor=cell.api, 
            Address=link_address, 
            TextToDisplay=display_text or link_address
        )
        logging.info(f"Đã thêm hyperlink tại ô '{cell_address}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi thêm hyperlink: {e}")
        return False

def group_rows(wb, sheet_name, start_row, end_row):
    """Nhóm các hàng lại với nhau (tạo outline)."""
    logging.debug(f"Bắt đầu nhóm hàng từ {start_row} đến {end_row} trên sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        sheet.range(f'{start_row}:{end_row}').rows.group()
        logging.info(f"Đã nhóm hàng từ {start_row} đến {end_row} thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi nhóm hàng: {e}")
        return False
        
# ======================================================================
# --- Nhóm 4: Dọn dẹp ---
# ======================================================================

def clear_range(wb, sheet_name, range_address, contents_only=True):
    """Xóa dữ liệu trong một vùng (chỉ nội dung hoặc cả định dạng)."""
    action = "nội dung" if contents_only else "toàn bộ"
    logging.debug(f"Bắt đầu xóa {action} của vùng '{range_address}' trên sheet '{sheet_name}'.")
    try:
        rng = wb.sheets[sheet_name].range(range_address)
        if contents_only:
            rng.clear_contents()
        else:
            rng.clear()
        logging.info(f"Đã xóa {action} của vùng '{range_address}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi xóa vùng '{range_address}': {e}")
        return False

