# Đường dẫn: excel_toolkit/utils/worksheet_ops.py
# Phiên bản 5.1 - Bổ sung hàm unhide_all_sheets
# Ngày cập nhật: 2025-09-12

import logging
import xlwings as xw

# ======================================================================
# --- Nhóm 1: Lấy thông tin & Trạng thái ---
# ======================================================================

def is_sheet_exist(wb, sheet_name):
    """Kiểm tra xem một sheet có tồn tại trong workbook hay không."""
    logging.debug(f"Bắt đầu kiểm tra sự tồn tại của sheet '{sheet_name}' trong workbook '{wb.name}'.")
    try:
        wb.sheets[sheet_name]
        logging.debug(f"  -> Kết quả: Sheet '{sheet_name}' tồn tại.")
        return True
    except:
        logging.debug(f"  -> Kết quả: Sheet '{sheet_name}' không tồn tại.")
        return False

def get_sheets_visibility(wb):
    """Lấy danh sách các sheet hiển thị và ẩn trong một workbook."""
    logging.debug(f"Bắt đầu lấy danh sách sheet ẩn/hiện trong workbook: {wb.name}")
    try:
        visible_sheets = [sheet.name for sheet in wb.sheets if sheet.api.Visible == -1]
        hidden_sheets = [sheet.name for sheet in wb.sheets if sheet.api.Visible == 0]
        logging.info(f"Đã lấy danh sách sheet: {len(visible_sheets)} hiển thị, {len(hidden_sheets)} ẩn.")
        return visible_sheets, hidden_sheets
    except Exception as e:
        logging.error(f"Lỗi khi lấy danh sách sheet ẩn/hiện: {e}")
        return [], []

def get_all_sheet_names(wb):
    """Trả về một danh sách chứa tên của tất cả các sheet."""
    logging.debug(f"Bắt đầu lấy tên tất cả các sheet trong workbook '{wb.name}'.")
    try:
        sheet_names = [sheet.name for sheet in wb.sheets]
        logging.info(f"Đã lấy thành công {len(sheet_names)} tên sheet.")
        return sheet_names
    except Exception as e:
        logging.error(f"Lỗi khi lấy tên các sheet: {e}")
        return []

def get_active_sheet_name(wb):
    """Lấy tên của sheet đang được kích hoạt (active)."""
    logging.debug(f"Bắt đầu lấy tên sheet đang active trong workbook '{wb.name}'.")
    try:
        active_sheet = wb.sheets.active
        logging.info(f"Sheet đang active là: '{active_sheet.name}'.")
        return active_sheet.name
    except Exception as e:
        logging.error(f"Lỗi khi lấy tên sheet active: {e}")
        return None

def count_visible_sheets(wb):
    """Đếm tổng số sheet hiển thị trong một workbook."""
    logging.debug(f"Bắt đầu đếm các sheet hiển thị trong workbook '{wb.name}'")
    try:
        count = sum(1 for sheet in wb.sheets if sheet.api.Visible == -1)
        logging.info(f"Tổng số sheet hiển thị: {count}")
        return count
    except Exception as e:
        logging.error(f"Lỗi khi đếm các sheet hiển thị: {e}")
        return 0

def count_hidden_sheets(wb):
    """Đếm tổng số sheet ẩn trong một workbook."""
    logging.debug(f"Bắt đầu đếm các sheet ẩn trong workbook '{wb.name}'.")
    try:
        count = sum(1 for sheet in wb.sheets if sheet.api.Visible == 0)
        logging.info(f"Tổng số sheet ẩn: {count}")
        return count
    except Exception as e:
        logging.error(f"Lỗi khi đếm các sheet ẩn: {e}")
        return 0

# ======================================================================
# --- Nhóm 2: Thêm, Sửa, Xóa Sheet ---
# ======================================================================

def add_sheet(wb, sheet_name, after=None, before=None):
    """Tạo một sheet mới, có thể tùy chọn vị trí."""
    logging.debug(f"Bắt đầu tạo sheet mới '{sheet_name}' trong workbook '{wb.name}'.")
    try:
        after_sheet = wb.sheets[after] if after else None
        before_sheet = wb.sheets[before] if before else None
        new_sheet = wb.sheets.add(name=sheet_name, after=after_sheet, before=before_sheet)
        logging.info(f"Đã tạo thành công sheet mới: '{new_sheet.name}'.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi tạo sheet mới '{sheet_name}': {e}")
        return False

def rename_sheet(wb, sheet_name, new_name):
    """Đổi tên một sheet trong workbook."""
    logging.debug(f"Bắt đầu đổi tên sheet '{sheet_name}' thành '{new_name}'.")
    try:
        wb.sheets[sheet_name].name = new_name
        logging.info(f"Đã đổi tên sheet thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi đổi tên sheet '{sheet_name}': {e}")
        return False

def delete_sheet(wb, sheet_name):
    """Xóa một sheet cụ thể khỏi workbook."""
    logging.debug(f"Bắt đầu xóa sheet '{sheet_name}'.")
    try:
        if len(wb.sheets) <= 1:
            logging.warning("Không thể xóa sheet cuối cùng trong workbook.")
            return False
        wb.sheets[sheet_name].delete()
        logging.info(f"Đã xóa sheet '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi xóa sheet '{sheet_name}': {e}")
        return False

def delete_hidden_sheets(wb):
    """Xóa tất cả các sheet bị ẩn trong một workbook."""
    logging.debug(f"Bắt đầu xóa các sheet ẩn trong workbook: {wb.name}")
    try:
        sheets_to_delete = [sheet for sheet in wb.sheets if sheet.api.Visible == 0]
        if len(wb.sheets) - len(sheets_to_delete) <= 0:
            logging.warning("Không thể xóa, vì sẽ không còn sheet nào được hiển thị.")
            return False, []
        
        deleted_sheets = []
        for sheet in sheets_to_delete:
            sheet_name = sheet.name
            sheet.delete()
            deleted_sheets.append(sheet_name)
            logging.debug(f"  -> Đã xóa sheet ẩn: {sheet_name}")
        logging.info(f"Đã xóa thành công {len(deleted_sheets)} sheet ẩn.")
        return True, deleted_sheets
    except Exception as e:
        logging.error(f"Lỗi khi xóa các sheet ẩn: {e}")
        return False, []

# ======================================================================
# --- Nhóm 3: Cấu trúc & Vị trí ---
# ======================================================================

def copy_sheet(wb, source_sheet_name, target_sheet_name, after=None, before=None):
    """Sao chép một sheet trong cùng một workbook."""
    logging.debug(f"Bắt đầu sao chép sheet '{source_sheet_name}' thành '{target_sheet_name}'.")
    try:
        source_sheet = wb.sheets[source_sheet_name]
        after_sheet = wb.sheets[after] if after else None
        before_sheet = wb.sheets[before] if before else None
        source_sheet.copy(name=target_sheet_name, after=after_sheet, before=before_sheet)
        logging.info(f"Đã sao chép sheet thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{source_sheet_name}' hoặc sheet tham chiếu vị trí.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi sao chép sheet: {e}")
        return False

def move_sheet(wb, sheet_name, after=None, before=None):
    """Di chuyển một sheet đến một vị trí mới."""
    logging.debug(f"Bắt đầu di chuyển sheet '{sheet_name}'.")
    try:
        sheet_to_move = wb.sheets[sheet_name]
        after_sheet = wb.sheets[after] if after else None
        before_sheet = wb.sheets[before] if before else None
        sheet_to_move.move(after=after_sheet, before=before_sheet)
        logging.info(f"Đã di chuyển sheet '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}' hoặc sheet tham chiếu vị trí.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi di chuyển sheet: {e}")
        return False

def activate_sheet(wb, sheet_name):
    """Kích hoạt (chuyển sang) một sheet cụ thể."""
    logging.debug(f"Bắt đầu kích hoạt sheet '{sheet_name}'.")
    try:
        wb.sheets[sheet_name].activate()
        logging.info(f"Đã kích hoạt sheet '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi kích hoạt sheet: {e}")
        return False

# ======================================================================
# --- Nhóm 4: Bảo vệ & An toàn ---
# ======================================================================

def protect_sheet(wb, sheet_name, password=''):
    """Khóa một sheet để ngăn chỉnh sửa, có thể tùy chọn mật khẩu."""
    logging.debug(f"Bắt đầu khóa sheet '{sheet_name}'.")
    try:
        wb.sheets[sheet_name].protect(password=password)
        logging.info(f"Đã khóa sheet '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi khóa sheet: {e}")
        return False

def unprotect_sheet(wb, sheet_name, password=''):
    """Mở khóa một sheet đã được bảo vệ."""
    logging.debug(f"Bắt đầu mở khóa sheet '{sheet_name}'.")
    try:
        wb.sheets[sheet_name].unprotect(password=password)
        logging.info(f"Đã mở khóa sheet '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi mở khóa sheet: {e}")
        return False

# ======================================================================
# --- Nhóm 5: Dọn dẹp & Trực quan ---
# ======================================================================

def clear_sheet(wb, sheet_name, contents_only=True):
    """Xóa dữ liệu trên sheet (chỉ nội dung hoặc cả định dạng)."""
    action = "nội dung" if contents_only else "toàn bộ (nội dung và định dạng)"
    logging.debug(f"Bắt đầu xóa {action} trên sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        if contents_only:
            sheet.clear_contents()
        else:
            sheet.clear()
        logging.info(f"Đã xóa {action} trên sheet '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi xóa sheet: {e}")
        return False

def set_sheet_visibility(wb, sheet_name, visible=True):
    """Ẩn hoặc hiện một sheet cụ thể."""
    action = "hiển thị" if visible else "ẩn"
    logging.debug(f"Bắt đầu {action} sheet '{sheet_name}'.")
    try:
        wb.sheets[sheet_name].api.Visible = -1 if visible else 0
        logging.info(f"Đã {action} sheet '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi {action} sheet: {e}")
        return False

def set_sheet_tab_color(wb, sheet_name, color_rgb):
    """Đổi màu của tab sheet."""
    logging.debug(f"Bắt đầu đổi màu tab sheet '{sheet_name}' thành {color_rgb}.")
    try:
        wb.sheets[sheet_name].api.Tab.Color = color_rgb[0] + (color_rgb[1] * 256) + (color_rgb[2] * 256 * 256)
        logging.info(f"Đã đổi màu tab sheet thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi đổi màu tab sheet: {e}")
        return False

def delete_all_comments(wb, sheet_name):
    """Xóa tất cả các ghi chú (comments) trên một sheet."""
    logging.debug(f"Bắt đầu xóa tất cả comments trên sheet '{sheet_name}'.")
    try:
        wb.sheets[sheet_name].api.Cells.ClearComments()
        logging.info("Đã xóa tất cả comments thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi xóa comments: {e}")
        return False

def remove_all_hyperlinks(wb, sheet_name):
    """Xóa tất cả siêu liên kết (hyperlinks) trên sheet nhưng giữ lại văn bản."""
    logging.debug(f"Bắt đầu xóa tất cả hyperlinks trên sheet '{sheet_name}'.")
    try:
        wb.sheets[sheet_name].api.Hyperlinks.Delete()
        logging.info("Đã xóa tất cả hyperlinks thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi xóa hyperlinks: {e}")
        return False

def unhide_all_sheets(wb):
    """
    Hiển thị tất cả các sheet đang bị ẩn trong workbook.
    """
    logging.debug(f"Bắt đầu hiển thị tất cả các sheet ẩn trong workbook '{wb.name}'.")
    unhidden_count = 0
    try:
        for sheet in wb.sheets:
            if sheet.api.Visible != -1: # -1 = xlSheetVisible
                sheet.api.Visible = -1
                logging.debug(f"  -> Đã hiển thị sheet: '{sheet.name}'")
                unhidden_count += 1
        
        if unhidden_count > 0:
            logging.info(f"Đã hiển thị thành công {unhidden_count} sheet.")
        else:
            logging.info("Không tìm thấy sheet nào đang bị ẩn.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi hiển thị các sheet ẩn: {e}")
        return False
        
# ======================================================================
# --- Nhóm 6: Bố cục & Hiển thị ---
# ======================================================================

def get_used_range_address(wb, sheet_name):
    """Trả về địa chỉ của vùng dữ liệu đã sử dụng trên sheet."""
    logging.debug(f"Bắt đầu lấy địa chỉ used_range cho sheet '{sheet_name}'.")
    try:
        address = wb.sheets[sheet_name].used_range.address
        logging.info(f"Địa chỉ used_range là: {address}")
        return address
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return None
    except Exception as e:
        logging.error(f"Lỗi khi lấy used_range: {e}")
        return None

def unfreeze_panes(wb, sheet_name):
    """Gỡ bỏ việc cố định hàng/cột (freeze panes)."""
    logging.debug(f"Bắt đầu gỡ bỏ freeze panes cho sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        if sheet.api.FreezePanes:
            sheet.api.FreezePanes = False
            logging.info("Đã gỡ bỏ freeze panes thành công.")
        else:
            logging.info("Không có freeze panes nào được thiết lập để gỡ bỏ.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi gỡ bỏ freeze panes: {e}")
        return False
        
def ungroup_all_rows(wb, sheet_name):
    """Gỡ bỏ tất cả các nhóm hàng (group/outline) trên một sheet."""
    logging.debug(f"Bắt đầu gỡ bỏ tất cả các nhóm hàng trên sheet '{sheet_name}'.")
    try:
        wb.sheets[sheet_name].api.Rows.Ungroup()
        logging.info("Đã gỡ bỏ tất cả các nhóm hàng thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi gỡ bỏ nhóm hàng: {e}")
        return False

# ======================================================================
# --- Nhóm 7: Chế độ xem Nâng cao ---
# ======================================================================

def set_zoom(wb, sheet_name, zoom_percentage=100):
    """Đặt mức độ phóng to/thu nhỏ cho một sheet."""
    logging.debug(f"Bắt đầu đặt zoom {zoom_percentage}% cho sheet '{sheet_name}'.")
    try:
        wb.sheets[sheet_name].api.Zoom = zoom_percentage
        logging.info("Đã đặt zoom thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi đặt zoom: {e}")
        return False

def toggle_gridlines(wb, sheet_name, show=True):
    """Bật hoặc tắt đường lưới (gridlines) trên một sheet."""
    action = "bật" if show else "tắt"
    logging.debug(f"Bắt đầu {action} gridlines cho sheet '{sheet_name}'.")
    try:
        wb.app.api.ActiveWindow.DisplayGridlines = show
        logging.info(f"Đã {action} gridlines thành công.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi {action} gridlines: {e}")
        return False

def toggle_headings(wb, sheet_name, show=True):
    """Bật hoặc tắt tiêu đề hàng và cột (A, B, C... và 1, 2, 3...)."""
    action = "bật" if show else "tắt"
    logging.debug(f"Bắt đầu {action} headings cho sheet '{sheet_name}'.")
    try:
        wb.app.api.ActiveWindow.DisplayHeadings = show
        logging.info(f"Đã {action} headings thành công.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi {action} headings: {e}")
        return False

# ======================================================================
# --- Nhóm 8: Tìm kiếm & Thay thế ---
# ======================================================================

def is_text_in_sheet(wb, sheet_name, text, exact_match=False):
    """Kiểm tra xem văn bản có tồn tại trong sheet hay không."""
    logging.debug(f"Bắt đầu kiểm tra sự tồn tại của '{text}' trong sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        look_at = xw.constants.LookAt.xlWhole if exact_match else xw.constants.LookAt.xlPart
        found_cell = sheet.api.Cells.Find(What=text, LookAt=look_at)
        
        if found_cell:
            logging.debug(f"  -> Kết quả: Tìm thấy '{text}'.")
            return True
        else:
            logging.debug(f"  -> Kết quả: Không tìm thấy '{text}'.")
            return False
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi tìm kiếm văn bản: {e}")
        return False

def find_all_in_sheet(wb, sheet_name, text, exact_match=False):
    """Tìm và trả về địa chỉ của tất cả các ô chứa văn bản."""
    logging.debug(f"Bắt đầu tìm tất cả các ô chứa '{text}' trong sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        look_at = xw.constants.LookAt.xlWhole if exact_match else xw.constants.LookAt.xlPart
        
        found_cell = sheet.api.Cells.Find(What=text, LookAt=look_at)
        if not found_cell:
            logging.info(f"Không tìm thấy ô nào chứa '{text}'.")
            return []

        addresses = []
        first_address = found_cell.Address
        addresses.append(first_address)
        
        while True:
            found_cell = sheet.api.Cells.FindNext(found_cell)
            if found_cell is None or found_cell.Address == first_address:
                break
            addresses.append(found_cell.Address)
            
        logging.info(f"Tìm thấy {len(addresses)} ô chứa '{text}': {addresses}")
        return addresses
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return []
    except Exception as e:
        logging.error(f"Lỗi khi tìm kiếm văn bản: {e}")
        return []

def replace_in_sheet(wb, sheet_name, search_text, replace_text, exact_match=False):
    """Tìm và thay thế tất cả các lần xuất hiện của văn bản trong sheet."""
    logging.debug(f"Bắt đầu thay thế '{search_text}' bằng '{replace_text}' trong sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        look_at = xw.constants.LookAt.xlWhole if exact_match else xw.constants.LookAt.xlPart
        
        # xlwings không có hàm replace trực tiếp, phải dùng API
        sheet.api.Cells.Replace(
            What=search_text,
            Replacement=replace_text,
            LookAt=look_at
        )
        logging.info("Hoàn tất quá trình thay thế.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi thay thế văn bản: {e}")
        return False

