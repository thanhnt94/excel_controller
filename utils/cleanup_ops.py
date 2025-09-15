# Đường dẫn: excel_toolkit/utils/cleanup_ops.py
# Phiên bản 2.2 - Khắc phục lỗi `xlwings.utils.col_str`
# Ngày cập nhật: 2025-09-15

import logging
import xlwings as xw

def _col_to_str(col_index):
    """Chuyển đổi chỉ số cột (số) thành ký tự cột (A, B, C...)."""
    string = ""
    while col_index > 0:
        col_index, remainder = divmod(col_index - 1, 26)
        string = chr(65 + remainder) + string
    return string

# ======================================================================
# --- Nhóm 1: Dọn dẹp Cấu trúc Workbook ---
# ======================================================================

def delete_external_links(wb):
    """
    Xóa tất cả các liên kết đến file Excel bên ngoài trong workbook.
    """
    logging.debug(f"Bắt đầu xóa liên kết ngoài cho workbook '{wb.name}'.")
    try:
        link_sources = wb.api.LinkSources(1)
        if link_sources:
            logging.info(f"Đã tìm thấy {len(link_sources)} liên kết ngoài. Bắt đầu ngắt liên kết...")
            for i in range(len(link_sources), 0, -1):
                try:
                    link_path = link_sources[i-1]
                    wb.api.BreakLink(link_path, 1)
                    logging.debug(f"  -> Đã ngắt liên kết đến: {link_path}")
                except Exception as break_err:
                    logging.warning(f"Không thể ngắt liên kết '{link_sources[i-1]}'. Lỗi: {break_err}")
            logging.info("Hoàn tất việc xóa liên kết ngoài.")
        else:
            logging.info("Không tìm thấy liên kết ngoài nào trong workbook.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi xóa liên kết ngoài: {e}")
        return False

def delete_defined_names(wb):
    """
    Xóa tất cả các 'Defined Names' trong workbook, bỏ qua thiết lập in.
    """
    logging.debug(f"Bắt đầu xóa 'Defined Names' cho workbook '{wb.name}'.")
    deleted_count, skipped_count = 0, 0
    try:
        for i in range(len(wb.api.Names), 0, -1):
            name = wb.api.Names(i)
            name_text = name.Name
            if "Print_Area" in name_text or "Print_Titles" in name_text:
                logging.debug(f"  -> Bỏ qua name thiết lập in: '{name_text}'")
                skipped_count += 1
                continue
            try:
                name.Delete()
                logging.debug(f"  -> Đã xóa name: '{name_text}'")
                deleted_count += 1
            except Exception as delete_err:
                logging.warning(f"Không thể xóa name '{name_text}'. Lỗi: {delete_err}")
        logging.info(f"Hoàn tất xóa 'Defined Names'. Đã xóa: {deleted_count}, Bỏ qua: {skipped_count}.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi xóa 'Defined Names': {e}")
        return False

# ======================================================================
# --- Nhóm 2: Tối ưu hóa & Bảo mật ---
# ======================================================================

def remove_personal_info(wb):
    """
    Xóa các thông tin cá nhân và siêu dữ liệu khỏi thuộc tính của file.
    """
    logging.debug(f"Bắt đầu xóa thông tin cá nhân khỏi workbook '{wb.name}'.")
    try:
        wb.api.RemoveDocumentInformation(10)
        logging.info("Đã xóa thành công thông tin cá nhân khỏi file.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi xóa thông tin cá nhân: {e}")
        return False

def clear_excess_cell_formatting(wb):
    """
    Xóa các định dạng ô không cần thiết nằm ngoài vùng dữ liệu đã sử dụng (used_range).
    """
    logging.debug(f"Bắt đầu xóa định dạng ô thừa cho workbook '{wb.name}'.")
    try:
        for sheet in wb.sheets:
            if sheet.api.Visible == -1:
                logging.debug(f"  -> Đang xử lý sheet: '{sheet.name}'")
                used_range = sheet.used_range
                last_row, last_col = used_range.last_cell.row, used_range.last_cell.column
                if last_row < sheet.api.Rows.Count:
                    range_to_clear_rows = sheet.range((last_row + 1, 1), (sheet.api.Rows.Count, last_col))
                    range_to_clear_rows.clear_formats()
                    logging.debug(f"    -> Đã xóa định dạng từ hàng {last_row + 1} trở xuống.")
                if last_col < sheet.api.Columns.Count:
                    range_to_clear_cols = sheet.range((1, last_col + 1), (sheet.api.Rows.Count, sheet.api.Columns.Count))
                    range_to_clear_cols.clear_formats()
                    # Sử dụng hàm trợ giúp mới
                    col_letter = _col_to_str(last_col + 1)
                    logging.debug(f"    -> Đã xóa định dạng từ cột {col_letter} trở đi.")
        logging.info("Hoàn tất việc xóa định dạng ô thừa.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi xóa định dạng ô thừa: {e}")
        return False

def refresh_and_clean_pivot_caches(wb):
    """
    Làm mới tất cả các Pivot Table và dọn dẹp cache của chúng để giảm dung lượng.
    """
    logging.debug(f"Bắt đầu làm mới và dọn dẹp Pivot Table caches cho workbook '{wb.name}'.")
    try:
        if not wb.api.PivotCaches().Count > 0:
            logging.info("Không tìm thấy Pivot Table cache nào trong workbook.")
            return True

        for cache in wb.api.PivotCaches():
            cache.SaveData = False
            cache.Refresh()
            logging.debug("  -> Đã làm mới và tắt SaveData cho một Pivot Cache.")
        
        logging.info("Đã làm mới và dọn dẹp thành công tất cả các Pivot Table cache.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi dọn dẹp Pivot Table cache: {e}")
        return False
