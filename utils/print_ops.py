# Đường dẫn: excel_toolkit/utils/print_ops.py
# Phiên bản 2.1 - Khắc phục lỗi `xlwings.utils.col_str`
# Ngày cập nhật: 2025-09-15

import logging
import xlwings as xw

# Các hằng số cho PageSetup (giúp code dễ đọc hơn)
A4_PAPER = 9
A3_PAPER = 8
PORTRAIT_ORIENTATION = 1
LANDSCAPE_ORIENTATION = 2

def _col_to_str(col_index):
    """Chuyển đổi chỉ số cột (số) thành ký tự cột (A, B, C...)."""
    string = ""
    while col_index > 0:
        col_index, remainder = divmod(col_index - 1, 26)
        string = chr(65 + remainder) + string
    return string

# ======================================================================
# --- Nhóm 1: Thiết lập Vùng in & Tiêu đề ---
# ======================================================================

def set_print_area(wb, sheet_name, print_range=None):
    """Thiết lập vùng in cho một sheet (toàn bộ vùng sử dụng hoặc một vùng cụ thể)."""
    logging.debug(f"Bắt đầu thiết lập vùng in cho sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        if print_range:
            sheet.api.PageSetup.PrintArea = print_range
            logging.info(f"Đã đặt vùng in cho '{sheet_name}' thành '{print_range}'.")
        else:
            used_range_address = sheet.used_range.address
            sheet.api.PageSetup.PrintArea = used_range_address
            logging.info(f"Đã đặt vùng in cho '{sheet_name}' thành toàn bộ vùng đã sử dụng: '{used_range_address}'.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi thiết lập vùng in cho '{sheet_name}': {e}")
        return False

def set_print_title_rows(wb, sheet_name, start_row, end_row):
    """Thiết lập các hàng tiêu đề sẽ lặp lại ở mỗi trang in."""
    logging.debug(f"Bắt đầu thiết lập hàng tiêu đề in từ {start_row} đến {end_row} cho sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        row_range = f"${start_row}:${end_row}"
        sheet.api.PageSetup.PrintTitleRows = row_range
        logging.info(f"Đã đặt hàng tiêu đề in cho '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi thiết lập hàng tiêu đề in cho '{sheet_name}': {e}")
        return False

def set_print_title_columns(wb, sheet_name, start_col, end_col):
    """Thiết lập các cột tiêu đề sẽ lặp lại ở mỗi trang in."""
    logging.debug(f"Bắt đầu thiết lập cột tiêu đề in từ {start_col} đến {end_col} cho sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        col_range = f"${_col_to_str(start_col)}:${_col_to_str(end_col)}"
        sheet.api.PageSetup.PrintTitleColumns = col_range
        logging.info(f"Đã đặt cột tiêu đề in cho '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi thiết lập cột tiêu đề in cho '{sheet_name}': {e}")
        return False

# ======================================================================
# --- Nhóm 2: Thiết lập Bố cục Trang ---
# ======================================================================

def set_page_orientation(wb, sheet_name, orientation=LANDSCAPE_ORIENTATION):
    """Đặt hướng trang in (dọc/ngang)."""
    orientation_text = "Ngang (Landscape)" if orientation == LANDSCAPE_ORIENTATION else "Dọc (Portrait)"
    logging.debug(f"Bắt đầu đặt hướng trang '{orientation_text}' cho sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        if orientation == PORTRAIT_ORIENTATION:
            sheet.api.PageSetup.Orientation = xw.constants.PageOrientation.xlPortrait
        elif orientation == LANDSCAPE_ORIENTATION:
            sheet.api.PageSetup.Orientation = xw.constants.PageOrientation.xlLandscape
        else:
            logging.error(f"Hướng trang không hợp lệ: {orientation}. Vui lòng sử dụng hằng số.")
            return False
        logging.info(f"Đã đặt hướng trang của '{sheet_name}' thành '{orientation_text}'.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi thiết lập hướng trang cho '{sheet_name}': {e}")
        return False

def set_fit_to_page(wb, sheet_name, fit_to_wide=1, fit_to_tall=False):
    """Thiết lập co giãn để vừa với trang in."""
    logging.debug(f"Bắt đầu thiết lập co giãn trang in cho sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        sheet.api.PageSetup.FitToPagesWide = fit_to_wide
        sheet.api.PageSetup.FitToPagesTall = fit_to_tall
        logging.info(f"Đã đặt '{sheet_name}' vừa với {fit_to_wide} trang rộng và {fit_to_tall} trang cao.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi thiết lập co giãn trang in cho '{sheet_name}': {e}")
        return False

def set_paper_size(wb, sheet_name, paper_size=A3_PAPER):
    """Thiết lập khổ giấy in (A3, A4, etc.)."""
    logging.debug(f"Bắt đầu thiết lập khổ giấy '{paper_size}' cho sheet '{sheet_name}'.")
    try:
        sheet = wb.sheets[sheet_name]
        sheet.api.PageSetup.PaperSize = paper_size
        logging.info(f"Đã đặt khổ giấy cho '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi thiết lập khổ giấy cho '{sheet_name}': {e}")
        return False

# ======================================================================
# --- Nhóm 3: Quản lý Header & Footer ---
# ======================================================================

def set_header_footer(wb, sheet_name, header_left="", header_center="", header_right="", footer_left="", footer_center="", footer_right=""):
    """Thiết lập nội dung cho Header và Footer."""
    logging.debug(f"Bắt đầu thiết lập header/footer cho sheet '{sheet_name}'.")
    try:
        page_setup = wb.sheets[sheet_name].api.PageSetup
        page_setup.LeftHeader = header_left
        page_setup.CenterHeader = header_center
        page_setup.RightHeader = header_right
        page_setup.LeftFooter = footer_left
        page_setup.CenterFooter = footer_center
        page_setup.RightFooter = footer_right
        logging.info(f"Đã thiết lập header/footer cho sheet '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi thiết lập header/footer: {e}")
        return False

# ======================================================================
# --- Nhóm 4: Các Tùy chọn In khác ---
# ======================================================================

def set_margins(wb, sheet_name, top=72, bottom=72, left=54, right=54, header=36, footer=36):
    """Thiết lập lề cho trang in (đơn vị: points, 72 points = 1 inch)."""
    logging.debug(f"Bắt đầu thiết lập lề cho sheet '{sheet_name}'.")
    try:
        page_setup = wb.sheets[sheet_name].api.PageSetup
        page_setup.TopMargin = top
        page_setup.BottomMargin = bottom
        page_setup.LeftMargin = left
        page_setup.RightMargin = right
        page_setup.HeaderMargin = header
        page_setup.FooterMargin = footer
        logging.info(f"Đã thiết lập lề cho sheet '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi thiết lập lề: {e}")
        return False

def toggle_print_options(wb, sheet_name, gridlines=False, headings=False, black_and_white=False):
    """Bật/tắt các tùy chọn in như đường lưới, tiêu đề, in đen trắng."""
    logging.debug(f"Bắt đầu thiết lập các tùy chọn in cho sheet '{sheet_name}'.")
    try:
        page_setup = wb.sheets[sheet_name].api.PageSetup
        page_setup.PrintGridlines = gridlines
        page_setup.PrintHeadings = headings
        page_setup.BlackAndWhite = black_and_white
        logging.info(f"Đã thiết lập các tùy chọn in cho sheet '{sheet_name}' thành công.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi thiết lập các tùy chọn in: {e}")
        return False

# ======================================================================
# --- Nhóm 5: Quy trình Tự động ---
# ======================================================================

def set_smart_print_settings(wb):
    """
    Áp dụng một bộ cài đặt in thông minh (A3, Ngang, co giãn theo chiều rộng)
    cho tất cả các sheet đang hiển thị trong workbook.
    """
    logging.debug(f"Bắt đầu áp dụng cài đặt in thông minh cho workbook '{wb.name}'.")
    try:
        visible_sheets, _ = xw.apps.active.books.active.sheets, [] # Tạm thời, sẽ thay bằng hàm visibility
        # Cần một hàm để lấy sheet hiển thị, giả sử ta đã có nó
        from . import worksheet_ops 
        visible_sheets, _ = worksheet_ops.get_sheets_visibility(wb)

        for sheet_name in visible_sheets:
            logging.debug(f"  -> Áp dụng cho sheet: '{sheet_name}'")
            sheet = wb.sheets[sheet_name]
            page_setup = sheet.api.PageSetup
            
            page_setup.PaperSize = A3_PAPER
            page_setup.Orientation = xw.constants.PageOrientation.xlLandscape
            page_setup.PrintArea = sheet.used_range.address
            page_setup.FitToPagesWide = 1
            page_setup.FitToPagesTall = False
        
        logging.info(f"Đã áp dụng cài đặt in thông minh cho {len(visible_sheets)} sheet thành công.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi áp dụng cài đặt in thông minh: {e}")
        return False
