# Đường dẫn: excel_toolkit/excel_controller.py
# Phiên bản: 4.9 - Sửa lỗi logic gọi engine Spire
# Ngày cập nhật: 2025-09-15

import logging
import xlwings as xw
import os 

from utils import (
    app_ops, cleanup_ops, convert_ops, data_ops, file_system_ops,
    print_ops, range_ops, shape_ops, worksheet_ops, 
    compressor_engine_pil, compressor_engine_spire
)

class ExcelController:
    """
    Lớp điều khiển trung tâm (Facade) cho framework Excel Toolkit.
    """
    def __init__(self, visible=False, optimize_performance=False):
        self.app = None
        self.workbook = None
        self.visible = visible
        self.optimize_performance = optimize_performance
        self.last_error = None
        
    def __enter__(self):
        try:
            self.app = xw.App(visible=self.visible)
            if self.optimize_performance:
                self.app.display_alerts = False
                self.app.screen_updating = False
            logging.info("Đã khởi tạo ứng dụng Excel.")
        except Exception as e:
            self.last_error = f"Lỗi khi khởi tạo ứng dụng Excel: {e}"
            logging.error(self.last_error)
            self.app = None
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.workbook:
            try:
                self.workbook.close()
            except Exception as e:
                logging.warning(f"Lỗi khi đóng workbook: {e}")
        
        if self.app:
            try:
                if self.optimize_performance:
                    self.app.display_alerts = True
                    self.app.screen_updating = True
                self.app.quit()
                logging.info("Đã thoát ứng dụng Excel.")
            except Exception as e:
                logging.error(f"Lỗi khi thoát ứng dụng Excel: {e}")

    # ======================================================================
    # --- 1. I/O Operations ---
    # ======================================================================

    def open_workbook(self, file_path, read_only=False, password="", ignore_read_only_recommended=True):
        if not file_system_ops.is_file_exist(file_path):
            self.last_error = f"Lỗi: File '{file_path}' không tồn tại."
            logging.error(self.last_error); return False
        if not self.app:
            self.last_error = "Lỗi: Ứng dụng Excel chưa được khởi tạo."
            logging.error(self.last_error); return False
        try:
            self.workbook = self.app.books.open(
                file_path, read_only=read_only, password=password,
                ignore_read_only_recommended=ignore_read_only_recommended
            )
            logging.info(f"Đã mở workbook thành công: '{os.path.basename(file_path)}'.")
            return True
        except Exception as e:
            self.last_error = f"Lỗi khi mở workbook '{file_path}': {e}"
            logging.error(self.last_error); self.workbook = None; return False
            
    def create_workbook(self, file_path=None):
        if not self.app:
            logging.error("Lỗi: Ứng dụng Excel chưa được khởi tạo."); return False
        try:
            self.workbook = self.app.books.add()
            if file_path:
                self.workbook.save(file_path)
                logging.info(f"Đã tạo và lưu workbook mới thành công tại '{file_path}'.")
            else:
                logging.info("Đã tạo workbook mới trong bộ nhớ.")
            return True
        except Exception as e:
            self.last_error = f"Lỗi khi tạo workbook mới: {e}"
            logging.error(self.last_error); self.workbook = None; return False

    def save_workbook(self, new_path=None):
        if not self.workbook:
            logging.warning("Không có workbook nào đang hoạt động để lưu."); return False
        try:
            save_path = new_path if new_path else self.workbook.fullname
            self.workbook.save(save_path)
            logging.info(f"Đã lưu workbook thành công tại '{save_path}'.")
            return True
        except Exception as e:
            self.last_error = f"Lỗi khi lưu workbook: {e}"
            logging.error(self.last_error); return False
            
    def close_workbook(self, save=True):
        if not self.workbook:
            logging.warning("Không có workbook nào đang hoạt động để đóng."); return False
        try:
            if save:
                self.workbook.save()
            self.workbook.close()
            logging.info("Đã đóng workbook."); self.workbook = None; return True
        except Exception as e:
            self.last_error = f"Lỗi khi đóng workbook: {e}"
            logging.error(self.last_error); return False
    
    # ======================================================================
    # --- 2. Worksheet Operations ---
    # ======================================================================

    def is_sheet_exist(self, sheet_name):
        return worksheet_ops.is_sheet_exist(self.workbook, sheet_name)
    def get_sheets_visibility(self):
        return worksheet_ops.get_sheets_visibility(self.workbook)
    def get_all_sheet_names(self):
        return worksheet_ops.get_all_sheet_names(self.workbook)
    def get_active_sheet_name(self):
        return worksheet_ops.get_active_sheet_name(self.workbook)
    def add_sheet(self, sheet_name, after=None, before=None):
        return worksheet_ops.add_sheet(self.workbook, sheet_name, after, before)
    def rename_sheet(self, sheet_name, new_name):
        return worksheet_ops.rename_sheet(self.workbook, sheet_name, new_name)
    def delete_sheet(self, sheet_name):
        return worksheet_ops.delete_sheet(self.workbook, sheet_name)
    def delete_hidden_sheets(self):
        return worksheet_ops.delete_hidden_sheets(self.workbook)
    def copy_sheet(self, source_sheet_name, target_sheet_name, after_sheet_name=None):
        return worksheet_ops.copy_sheet(self.workbook, source_sheet_name, target_sheet_name, after_sheet_name)
    def move_sheet(self, sheet_name, after=None, before=None):
        return worksheet_ops.move_sheet(self.workbook, sheet_name, after, before)
    def activate_sheet(self, sheet_name):
        return worksheet_ops.activate_sheet(self.workbook, sheet_name)
    def protect_sheet(self, sheet_name, password=''):
        return worksheet_ops.protect_sheet(self.workbook, sheet_name, password)
    def unprotect_sheet(self, sheet_name, password=''):
        return worksheet_ops.unprotect_sheet(self.workbook, sheet_name, password)
    def clear_sheet(self, sheet_name, contents_only=True):
        return worksheet_ops.clear_sheet(self.workbook, sheet_name, contents_only)
    def set_sheet_visibility(self, sheet_name, visible=True):
        return worksheet_ops.set_sheet_visibility(self.workbook, sheet_name, visible)
    def get_used_range_address(self, sheet_name):
        return worksheet_ops.get_used_range_address(self.workbook, sheet_name)
    def unfreeze_panes(self, sheet_name):
        return worksheet_ops.unfreeze_panes(self.workbook, sheet_name)
    def ungroup_all_rows(self, sheet_name):
        return worksheet_ops.ungroup_all_rows(self.workbook, sheet_name)
    def set_zoom(self, sheet_name, zoom_percentage=100):
        return worksheet_ops.set_zoom(self.workbook, sheet_name, zoom_percentage)
    def is_text_in_sheet(self, sheet_name, text, exact_match=False):
        return worksheet_ops.is_text_in_sheet(self.workbook, sheet_name, text, exact_match)
    def find_all_in_sheet(self, sheet_name, text, exact_match=False):
        return worksheet_ops.find_all_in_sheet(self.workbook, sheet_name, text, exact_match)
    def replace_in_sheet(self, sheet_name, search_text, replace_text, exact_match=False):
        return worksheet_ops.replace_in_sheet(self.workbook, sheet_name, search_text, replace_text, exact_match)
    def unhide_all_sheets(self):
        return worksheet_ops.unhide_all_sheets(self.workbook)

    # ======================================================================
    # --- 3. Range Operations ---
    # ======================================================================

    def get_cell_value(self, sheet_name, cell_address):
        return range_ops.get_cell_value(self.workbook, sheet_name, cell_address)
    def set_cell_value(self, sheet_name, cell_address, value):
        return range_ops.set_cell_value(self.workbook, sheet_name, cell_address, value)
    def get_range_values(self, sheet_name, range_address):
        return range_ops.get_range_values(self.workbook, sheet_name, range_address)
    def set_range_values(self, sheet_name, start_cell, values):
        return range_ops.set_range_values(self.workbook, sheet_name, start_cell, values)
    def get_last_row(self, sheet_name, column='A'):
        return range_ops.get_last_row(self.workbook, sheet_name, column)
    def format_range(self, sheet_name, range_address, format_properties):
        return range_ops.format_range(self.workbook, sheet_name, range_address, format_properties)
    def autofit_columns(self, sheet_name, range_address=None):
        return range_ops.autofit_columns(self.workbook, sheet_name, range_address)
    def freeze_panes(self, sheet_name, cell_address='B2'):
        return range_ops.freeze_panes(self.workbook, sheet_name, cell_address)
    def add_comment(self, sheet_name, cell_address, text):
        return range_ops.add_comment(self.workbook, sheet_name, cell_address, text)
    def group_rows(self, sheet_name, start_row, end_row):
        return range_ops.group_rows(self.workbook, sheet_name, start_row, end_row)

    # ======================================================================
    # --- 4. Shape Operations ---
    # ======================================================================
    
    def is_shape_exist(self, sheet_name, shape_name):
        return shape_ops.is_shape_exist(self.workbook, sheet_name, shape_name)
    def get_all_shape_names(self, sheet_name):
        return shape_ops.get_all_shape_names(self.workbook, sheet_name)
    def add_textbox(self, sheet_name, text, top, left, width, height, format_properties=None):
        return shape_ops.add_textbox(self.workbook, sheet_name, text, top, left, width, height, format_properties)
    def add_picture(self, sheet_name, image_path, top, left, width=None, height=None, name=None):
        return shape_ops.add_picture(self.workbook, sheet_name, image_path, top, left, width, height, name)
    def delete_shape(self, sheet_name, shape_name):
        return shape_ops.delete_shape(self.workbook, sheet_name, shape_name)
    
    # Hàm nén ảnh tổng hợp, cho phép chọn engine
    def compress_all_images(self, file_path, engine='pil', quality=70):
        if engine == 'pil':
            logging.info("Sử dụng engine 'Pillow' để nén ảnh.")
            # Pillow engine cần workbook object
            return compressor_engine_pil.compress_images(self.workbook, quality=quality)
        elif engine == 'spire':
            logging.info("Sử dụng engine 'Spire' để nén ảnh.")
            # Spire engine cần đường dẫn file
            # SỬA LỖI: Chỉ truyền một tham số đường dẫn
            return compressor_engine_spire.compress_images(file_path, max_size_kb=quality)
        else:
            logging.error(f"Engine nén ảnh '{engine}' không hợp lệ. Vui lòng chọn 'pil' hoặc 'spire'.")
            return False
            
    # ======================================================================
    # --- 5. Cleanup Operations ---
    # ======================================================================

    def delete_external_links(self):
        return cleanup_ops.delete_external_links(self.workbook)
    def delete_defined_names(self):
        return cleanup_ops.delete_defined_names(self.workbook)
    def remove_personal_info(self):
        return cleanup_ops.remove_personal_info(self.workbook)
    def clear_excess_cell_formatting(self):
        return cleanup_ops.clear_excess_cell_formatting(self.workbook)
    def refresh_and_clean_pivot_caches(self):
        return cleanup_ops.refresh_and_clean_pivot_caches(self.workbook)

    # ======================================================================
    # --- 6. Print Operations ---
    # ======================================================================
    
    def set_print_area(self, sheet_name, print_range=None):
        return print_ops.set_print_area(self.workbook, sheet_name, print_range)
    def set_print_title_rows(self, sheet_name, start_row, end_row):
        return print_ops.set_print_title_rows(self.workbook, sheet_name, start_row, end_row)
    def set_page_orientation(self, sheet_name, orientation):
        return print_ops.set_page_orientation(self.workbook, sheet_name, orientation)
    def set_fit_to_page(self, sheet_name, fit_to_wide=1, fit_to_tall=False):
        return print_ops.set_fit_to_page(self.workbook, sheet_name, fit_to_wide, fit_to_tall)
    def set_smart_print_settings(self):
        return print_ops.set_smart_print_settings(self.workbook)
        
    # ======================================================================
    # --- 7. Convert Operations ---
    # ======================================================================
    
    def excel_to_pdf(self, output_path):
        return convert_ops.excel_to_pdf(self.workbook, output_path)
    def sheet_to_pdf(self, sheet_name, output_path):
        return convert_ops.sheet_to_pdf(self.workbook, sheet_name, output_path)
    def save_sheet_as_csv(self, sheet_name, output_path):
        return convert_ops.save_sheet_as_csv(self.workbook, sheet_name, output_path)

