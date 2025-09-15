# Đường dẫn: excel_toolkit/utils/convert_ops.py
# Phiên bản 2.0 - Tái cấu trúc và bổ sung các hàm chuyển đổi đa năng
# Ngày cập nhật: 2025-09-12

import logging
import xlwings as xw
import os

# ======================================================================
# --- Nhóm 1: Chuyển đổi sang PDF ---
# ======================================================================

def workbook_to_pdf(wb, output_path):
    """
    Chuyển đổi toàn bộ workbook sang một file PDF duy nhất.
    Hàm này hoạt động với một workbook object đã mở.
    """
    logging.debug(f"Bắt đầu chuyển đổi workbook '{wb.name}' sang PDF tại '{output_path}'.")
    try:
        wb.to_pdf(output_path, include=None) # include=None để export tất cả các sheet
        logging.info(f"Đã chuyển đổi thành công workbook sang PDF: '{os.path.basename(output_path)}'.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi chuyển đổi workbook sang PDF: {e}")
        return False

def sheet_to_pdf(wb, sheet_name, output_path):
    """
    Chỉ chuyển đổi một sheet cụ thể sang định dạng PDF.
    """
    logging.debug(f"Bắt đầu chuyển đổi sheet '{sheet_name}' sang PDF tại '{output_path}'.")
    try:
        sheet_to_export = wb.sheets[sheet_name]
        sheet_to_export.to_pdf(output_path)
        logging.info(f"Đã chuyển đổi thành công sheet '{sheet_name}' sang PDF.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi chuyển đổi sheet sang PDF: {e}")
        return False

# ======================================================================
# --- Nhóm 2: Xuất dữ liệu sang các định dạng khác ---
# ======================================================================

def sheet_to_csv(wb, sheet_name, output_path, encoding='utf-8-sig'):
    """
    Lưu một sheet cụ thể dưới dạng file CSV.
    Mặc định sử dụng encoding 'utf-8-sig' để hỗ trợ tốt tiếng Việt.
    """
    logging.debug(f"Bắt đầu xuất sheet '{sheet_name}' sang file CSV tại '{output_path}'.")
    try:
        # Sử dụng pandas để thực hiện việc này một cách hiệu quả
        import pandas as pd
        sheet = wb.sheets[sheet_name]
        df = sheet.used_range.options(pd.DataFrame, index=False).value
        df.to_csv(output_path, index=False, encoding=encoding)
        logging.info(f"Đã xuất thành công sheet '{sheet_name}' sang CSV.")
        return True
    except ImportError:
        logging.error("Lỗi: Cần cài đặt thư viện 'pandas' để sử dụng chức năng này. (pip install pandas)")
        return False
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi xuất sheet sang CSV: {e}")
        return False

def range_to_image(wb, sheet_name, range_address, output_path):
    """
    Chụp một vùng dữ liệu trên sheet và lưu nó thành một file ảnh (PNG, JPG, ...).
    """
    logging.debug(f"Bắt đầu xuất vùng '{range_address}' trên sheet '{sheet_name}' sang ảnh tại '{output_path}'.")
    try:
        sheet = wb.sheets[sheet_name]
        range_obj = sheet.range(range_address)
        
        # Sao chép vùng dữ liệu như một ảnh
        range_obj.api.CopyPicture(Appearance=1, Format=2) # 1=xlScreen, 2=xlBitmap
        
        # Tạo một biểu đồ tạm thời để dán ảnh vào
        chart = sheet.charts.add()
        chart.api.Paste()
        
        # Xuất biểu đồ (chứa ảnh) ra file và xóa biểu đồ tạm
        chart.api.Export(output_path)
        chart.delete()
        
        logging.info(f"Đã xuất thành công vùng '{range_address}' sang ảnh.")
        return True
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi xuất vùng sang ảnh: {e}")
        return False
