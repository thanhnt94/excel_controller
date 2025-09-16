# Kịch bản nâng cao để đóng gói ứng dụng Excel Toolkit.
# Yêu cầu: pip install pyinstaller
# Cách dùng: Đặt file này vào thư mục gốc của dự án và chạy lệnh: python build_exe.py

import PyInstaller.__main__
import os
import shutil

# --- CẤU HÌNH CHÍNH ---
# Tự động lấy đường dẫn của thư mục chứa kịch bản này
script_dir = os.path.dirname(os.path.abspath(__file__))

# Tên file chính và tên file EXE đầu ra
main_script = os.path.join(script_dir, "main.py")
exe_name = 'Excel File Batch Processing Toolkit'

# Đường dẫn đến file icon (cần tạo thư mục assets và đặt file app_icon.ico vào đó)
icon_path = os.path.join(script_dir, 'assets', 'app_icon.ico') 

# --- CẤU HÌNH ĐƯỜNG DẪN ĐẦU RA (TƯƠNG ĐỐI) ---
# Tự động tạo các thư mục build và dist trong thư mục dự án
output_dir = os.path.join(script_dir, 'dist')
build_temp_dir = os.path.join(script_dir, 'build')

# (Tùy chọn) Đường dẫn đến thư mục chứa UPX để nén file EXE
upx_dir_path = 'C:/upx' 

# --- DANH SÁCH CÁC THƯ VIỆN CẦN LOẠI BỎ ---
# Giữ lại danh sách tối ưu hóa của bạn, loại bỏ các thư viện không cần thiết
modules_to_exclude = [
    # Web Frameworks & Scraping
    'Django', 'Flask', 'fastapi', 'uvicorn', 'asgiref', 'starlette', 'h11',
    'Scrapy', 'beautifulsoup4', 'pyquery', 'w3lib', 'tldextract', 'url_normalize', 
    'Twisted', 'zope.interface',
    
    # Data Science, Plotting, and Heavy Math
    'matplotlib', 'scipy', 'scikit-image', 'kiwisolver', 'contourpy', 'cycler',
    'fonttools', 'nltk', 'sympy', 'mpmath', 'networkx', 'joblib',

    # OCR and Deep Learning
    'easyocr', 'torch', 'torchvision', 'opencv-python', 'opencv-python-headless', 'ninja',
    
    # Alternative GUI / Automation Libraries
    'PyAutoGUI', 'PyMsgBox', 'PyRect', 'PyScreeze', 'MouseInfo', 'pytweening',

    # Browser Automation
    'playwright', 'pyppeteer', 'pyee', 'websockets',

    # Testing Frameworks
    'pytest', 'pluggy', 'iniconfig',

    # Other misc libraries not used in this project
    'pydantic', 'cattrs', 'defusedxml', 'imageio', 'jmespath', 'pyOpenSSL',
    'python-bidi', 'python-dotenv', 'xlrd', 'xlsxwriter', 'PyYAML'
]

# --- Dọn dẹp các bản build cũ ---
print("--- Dọn dẹp các thư mục build cũ... ---")
try:
    if os.path.isdir(build_temp_dir):
        shutil.rmtree(build_temp_dir)
    if os.path.isdir(output_dir):
        shutil.rmtree(output_dir)
    spec_file = os.path.join(script_dir, f'{exe_name}.spec')
    if os.path.isfile(spec_file):
        os.remove(spec_file)
    print("Dọn dẹp hoàn tất.")
except Exception as e:
    print(f"Lỗi khi dọn dẹp: {e}")


# --- Các tùy chọn cho PyInstaller ---
pyinstaller_options = [
    f'--name={exe_name}',
    '--onedir',      # Đóng gói thành một thư mục (ổn định hơn --onefile)
    '--windowed',    # Chạy dưới dạng ứng dụng cửa sổ, không có console
    f'--distpath={output_dir}',
    f'--workpath={build_temp_dir}',
    # --- Thêm các thư viện ẩn và tài nguyên ---
    '--collect-submodules=customtkinter',
    '--collect-submodules=spire' # Đảm bảo Spire.XLS được đóng gói đúng
]

# Thêm tùy chọn icon nếu file tồn tại
if os.path.isfile(icon_path):
    pyinstaller_options.append(f'--icon={icon_path}')
else:
    print(f"\n!!! CẢNH BÁO: Không tìm thấy file icon tại '{icon_path}'. Bỏ qua tùy chọn icon.")

# Thêm các module cần loại bỏ vào câu lệnh
for module in modules_to_exclude:
    pyinstaller_options.append(f'--exclude-module={module}')

# Kết hợp các tùy chọn với file kịch bản chính
full_command = pyinstaller_options + [main_script]

# --- CHẠY ĐÓNG GÓI ---
if __name__ == '__main__':
    print("\n--- Bắt đầu quá trình đóng gói với PyInstaller ---")
    
    # Kiểm tra và thêm tùy chọn UPX
    if os.path.isdir(upx_dir_path):
        print(f"Tìm thấy UPX tại: '{upx_dir_path}'. Sẽ sử dụng để nén file.")
        full_command.append(f'--upx-dir={upx_dir_path}')
    else:
        print(f"\n!!! CẢNH BÁO: Không tìm thấy thư mục UPX tại '{upx_dir_path}'.")
        print("Bỏ qua bước nén file. File EXE sẽ có dung lượng lớn hơn.")

    print(f"\nLệnh thực thi: pyinstaller {' '.join(full_command)}")
    
    try:
        # Tạo các thư mục đầu ra nếu chưa có
        os.makedirs(build_temp_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)
        
        # Chạy PyInstaller
        PyInstaller.__main__.run(full_command)
        
        print("\n--- QUÁ TRÌNH ĐÓNG GÓI HOÀN TẤT! ---")
        print(f"=> File thực thi của bạn nằm trong thư mục: {os.path.abspath(output_dir)}")
    except Exception as e:
        print("\n--- !!! CÓ LỖI XẢY RA TRONG QUÁ TRÌNH ĐÓNG GÓI !!! ---")
        print(f"Lỗi: {e}")
