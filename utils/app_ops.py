# Đường dẫn: excel_toolkit/utils/app_ops.py
# Phiên bản 2.0 - Tái cấu trúc và bổ sung các hàm tiện ích
# Ngày cập nhật: 2025-09-12

import logging
import psutil
import pygetwindow as gw
import win32process
import subprocess

# ======================================================================
# --- Nhóm 1: Kiểm tra trạng thái ứng dụng ---
# ======================================================================

def is_excel_running():
    """
    Kiểm tra xem có bất kỳ tiến trình Excel nào đang chạy hay không.
    """
    logging.debug("Kiểm tra xem có tiến trình EXCEL.EXE nào đang chạy không.")
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'EXCEL.EXE':
            logging.info("Đã tìm thấy ít nhất một tiến trình Excel đang chạy.")
            return True
    logging.info("Không tìm thấy tiến trình Excel nào đang chạy.")
    return False

# ======================================================================
# --- Nhóm 2: Thao tác đóng ứng dụng ---
# ======================================================================

def excel_force_close():
    """
    Buộc đóng tất cả các tiến trình Excel đang chạy bằng taskkill.
    Đây là phương pháp mạnh nhất để dọn dẹp môi trường.
    """
    logging.debug("Bắt đầu quá trình buộc đóng tất cả các tiến trình Excel.")
    try:
        # Dùng capture_output để ẩn output của taskkill khỏi console
        result = subprocess.run(["taskkill", "/f", "/im", "excel.exe"], check=True, capture_output=True, text=True)
        if "SUCCESS" in result.stdout:
            logging.info("Tất cả các tiến trình Excel đã được buộc đóng thành công.")
        else:
            # Trường hợp taskkill chạy nhưng không tìm thấy tiến trình nào
            logging.info("Không tìm thấy tiến trình Excel nào đang chạy để buộc đóng.")
        return True
    except subprocess.CalledProcessError:
        # Lỗi này xảy ra khi không tìm thấy tiến trình nào, không phải là lỗi nghiêm trọng
        logging.info("Không tìm thấy tiến trình Excel nào đang chạy để buộc đóng.")
        return True
    except FileNotFoundError:
        logging.error("Lỗi: Lệnh 'taskkill' không tồn tại. Vui lòng kiểm tra PATH hệ thống.")
        return False
    except Exception as e:
        logging.error(f"Lỗi không xác định khi buộc đóng Excel: {e}")
        return False

def excel_hidden_close():
    """
    Đóng tất cả các tiến trình Excel đang chạy ẩn (không có cửa sổ hiển thị).
    Hữu ích để dọn dẹp các tiến trình zombie mà không ảnh hưởng đến file người dùng đang mở.
    """
    logging.debug("Bắt đầu quá trình đóng các tiến trình Excel chạy ẩn.")
    
    def is_window_visible_for_pid(pid):
        """Kiểm tra xem một PID có cửa sổ Excel nào đang hiển thị không."""
        try:
            for window in gw.getWindowsWithTitle('Excel'):
                # Đảm bảo cửa sổ thực sự là một cửa sổ (có handle)
                if window._hWnd:
                    _, window_pid = win32process.GetWindowThreadProcessId(window._hWnd)
                    if window_pid == pid:
                        return True
        except Exception:
            # Bỏ qua nếu có lỗi khi tương tác với cửa sổ
            pass
        return False

    terminated_count = 0
    try:
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] == 'EXCEL.EXE':
                pid = proc.info['pid']
                if not is_window_visible_for_pid(pid):
                    try:
                        proc.terminate()
                        logging.info(f"Đã đóng tiến trình Excel ẩn (PID: {pid}).")
                        terminated_count += 1
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        logging.warning(f"Không thể đóng tiến trình Excel ẩn (PID: {pid}).")
        
        if terminated_count > 0:
            logging.info(f"Hoàn tất. Đã đóng thành công {terminated_count} tiến trình Excel ẩn.")
        else:
            logging.info("Không tìm thấy tiến trình Excel ẩn nào để đóng.")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi tìm và đóng tiến trình Excel ẩn: {e}")
        return False
