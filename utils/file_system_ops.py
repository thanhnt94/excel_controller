# Đường dẫn: excel_toolkit/utils/file_system_ops.py
# Phiên bản 2.1 - Sửa lỗi tên hàm cho nhất quán
# Ngày cập nhật: 2025-09-12

import logging
import os
import shutil
import stat

# ======================================================================
# --- Nhóm 1: Kiểm tra Trạng thái ---
# ======================================================================

def is_file_exist(file_path):
    """
    Kiểm tra xem một file có tồn tại hay không.
    (Tên hàm đã được sửa lại cho đúng)
    """
    logging.debug(f"Đang kiểm tra sự tồn tại của file: '{file_path}'")
    return os.path.isfile(file_path)

def is_folder_exist(folder_path):
    """
    Kiểm tra xem một thư mục có tồn tại hay không.
    (Tên hàm đã được sửa lại cho đúng)
    """
    logging.debug(f"Đang kiểm tra sự tồn tại của thư mục: '{folder_path}'")
    return os.path.isdir(folder_path)

# ======================================================================
# --- Nhóm 2: Thao tác File & Thư mục ---
# ======================================================================

def create_folder(folder_path):
    """
    Tạo một thư mục. Nếu các thư mục cha chưa tồn tại, chúng sẽ được tạo luôn.
    """
    logging.debug(f"Đang chuẩn bị tạo thư mục: '{folder_path}'")
    try:
        os.makedirs(folder_path, exist_ok=True)
        logging.info(f"Thư mục đã được tạo (hoặc đã tồn tại): '{folder_path}'")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi tạo thư mục '{folder_path}': {e}")
        return False

def delete_file(file_path):
    """
    Xóa một file một cách an toàn.
    """
    logging.debug(f"Đang chuẩn bị xóa file: '{file_path}'")
    try:
        if not is_file_exist(file_path):
            logging.warning(f"File không tồn tại, không thể xóa: '{file_path}'")
            return False
        os.remove(file_path)
        logging.info(f"Đã xóa file thành công: '{file_path}'")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi xóa file '{file_path}': {e}")
        return False

def delete_folder(folder_path):
    """
    Xóa một thư mục và toàn bộ nội dung bên trong nó một cách an toàn.
    """
    logging.debug(f"Đang chuẩn bị xóa thư mục: '{folder_path}'")
    try:
        if not is_folder_exist(folder_path):
            logging.warning(f"Thư mục không tồn tại, không thể xóa: '{folder_path}'")
            return False
        
        # Xử lý lỗi cho các file read-only trong Windows
        def on_rm_error(func, path, exc_info):
            os.chmod(path, stat.S_IWRITE)
            os.unlink(path)

        shutil.rmtree(folder_path, onerror=on_rm_error)
        logging.info(f"Đã xóa thư mục và nội dung thành công: '{folder_path}'")
        return True
    except Exception as e:
        logging.error(f"Lỗi khi xóa thư mục '{folder_path}': {e}")
        return False

# ======================================================================
# --- Nhóm 3: Lấy Thông tin & Duyệt file ---
# ======================================================================

def get_files_path(folder_path, file_extensions=None, include_subfolders=False):
    """
    Lấy danh sách các đường dẫn tuyệt đối của các file trong một thư mục.
    """
    logging.debug(f"Bắt đầu lấy đường dẫn file từ '{folder_path}'.")
    file_list = []
    if not is_folder_exist(folder_path):
        logging.error(f"Đường dẫn thư mục '{folder_path}' không tồn tại.")
        return []

    try:
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file_extensions:
                    # Chuyển đuôi file và đuôi trong danh sách về chữ thường để so sánh
                    if file.lower().endswith(tuple(ext.lower() for ext in file_extensions)):
                        file_list.append(os.path.join(root, file))
                else:
                    file_list.append(os.path.join(root, file))
            if not include_subfolders:
                break
        logging.info(f"Đã tìm thấy {len(file_list)} file phù hợp trong '{folder_path}'.")
        return file_list
    except Exception as e:
        logging.error(f"Lỗi khi lấy danh sách file: {e}")
        return []

def get_file_properties(file_path):
    """
    Lấy các thuộc tính của một file (kích thước, ngày tạo, ngày sửa đổi).
    """
    logging.debug(f"Đang lấy thuộc tính của file: '{file_path}'")
    if not is_file_exist(file_path):
        logging.error(f"File không tồn tại: '{file_path}'")
        return None
    try:
        stat_info = os.stat(file_path)
        properties = {
            'size_bytes': stat_info.st_size,
            'creation_time': stat_info.st_ctime,
            'modified_time': stat_info.st_mtime
        }
        logging.info(f"Đã lấy thuộc tính của file '{os.path.basename(file_path)}' thành công.")
        return properties
    except Exception as e:
        logging.error(f"Lỗi khi lấy thuộc tính file '{file_path}': {e}")
        return None

