# Đường dẫn: excel_toolkit/utils/image_compressor_spire_api.py
# Phiên bản 1.1 - Cập nhật để nhận tham số chất lượng
# Ngày cập nhật: 2025-09-15

from spire.xls import *
from spire.xls.common import *
import os
from PIL import Image
import io
import tempfile
import shutil
import uuid
import win32com.client
import pythoncom
import logging

def _optimize_image(input_path, output_path, max_size_kb=300):
    """
    Tối ưu hóa kích thước hình ảnh, đảm bảo không vượt quá kích thước chỉ định.
    """
    try:
        with Image.open(input_path) as img:
            max_width = 800
            max_height = 600
            width, height = img.size
            
            if width > max_width or height > max_height:
                ratio = min(max_width/width, max_height/height)
                new_size = (int(width*ratio), int(height*ratio))
                img = img.resize(new_size, Image.Resampling.LANCZOS)
            
            quality = 70
            if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                img.save(output_path, format='PNG', optimize=True, compression_level=9)
            else:
                img = img.convert('RGB')
                while quality > 10:
                    buffer = io.BytesIO()
                    img.save(buffer, format='JPEG', quality=quality, optimize=True, progressive=True)
                    if buffer.tell() <= max_size_kb * 1024 or quality <= 10:
                        img.save(output_path, format='JPEG', quality=quality, optimize=True, progressive=True)
                        break
                    quality -= 10
            
            return True
    except Exception as e:
        logging.error(f"Lỗi khi tối ưu hóa hình ảnh: {str(e)}")
        return False

def compress_excel_with_spire_api(input_file, output_file, max_size_kb=300):
    """
    Nén file Excel bằng cách trích xuất và tối ưu hóa hình ảnh với thư viện Spire.Xls.
    """
    logging.info("Bắt đầu nén ảnh bằng engine Spire.Xls...")
    temp_dir = tempfile.mkdtemp()
    compressed_dir = os.path.join(temp_dir, "compressed")
    os.makedirs(compressed_dir, exist_ok=True)
    
    try:
        workbook = Workbook()
        workbook.LoadFromFile(input_file)
        
        images_to_replace = []
        
        for sheet_index in range(workbook.Worksheets.Count):
            sheet = workbook.Worksheets[sheet_index]
            pic_count = sheet.Pictures.Count
            
            if pic_count > 0:
                logging.debug(f"  -> Đã tìm thấy {pic_count} ảnh trong sheet '{sheet.Name}'.")
            
            for i in range(pic_count):
                try:
                    pic = sheet.Pictures[i]
                    temp_filename = f"excel_img_{uuid.uuid4()}"
                    img_path = os.path.join(temp_dir, f"{temp_filename}.png")
                    
                    pic.Picture.Save(img_path)
                    
                    if os.path.exists(img_path) and os.path.getsize(img_path) > 0:
                        output_path = os.path.join(compressed_dir, f"compressed_{temp_filename}.png")
                        
                        if _optimize_image(img_path, output_path, max_size_kb):
                            image_info = {
                                'compressed_path': output_path,
                                'sheet_name': sheet.Name,
                                'left': pic.Left,
                                'top': pic.Top,
                                'width': pic.Width,
                                'height': pic.Height
                            }
                            images_to_replace.append(image_info)
                            
                            original_size_kb = os.path.getsize(img_path) / 1024
                            compressed_size_kb = os.path.getsize(output_path) / 1024
                            logging.info(f"    -> Đã nén ảnh thành công: {original_size_kb:.1f}KB -> {compressed_size_kb:.1f}KB")
                            
                except Exception as e:
                    logging.warning(f"Lỗi khi xử lý ảnh trong sheet '{sheet.Name}': {str(e)}")
                    continue
        
        if images_to_replace:
            logging.info(f"Đã tối ưu hóa {len(images_to_replace)} ảnh. Bắt đầu thay thế...")
            
            shutil.copy2(input_file, output_file)
            
            pythoncom.CoInitialize()
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                
                workbook_win32 = excel.Workbooks.Open(os.path.abspath(output_file))
                
                for img_info in images_to_replace:
                    sheet_name = img_info['sheet_name']
                    sheet = workbook_win32.Sheets(sheet_name)
                    
                    try:
                        for shape in sheet.Shapes:
                            if (abs(shape.Left - img_info['left']) < 5 and 
                                abs(shape.Top - img_info['top']) < 5 and 
                                shape.Type == 13): # msoPicture
                                shape.Delete()
                                logging.debug(f"    -> Đã xóa ảnh gốc trên sheet '{sheet_name}'.")
                                break
                    except Exception as e:
                        logging.warning(f"Không thể xóa ảnh gốc: {str(e)}")
                    
                    sheet.Shapes.AddPicture(
                        Filename=img_info['compressed_path'],
                        LinkToFile=False,
                        SaveWithDocument=True,
                        Left=img_info['left'],
                        Top=img_info['top'],
                        Width=img_info['width'],
                        Height=img_info['height']
                    )
                    logging.debug(f"    -> Đã chèn ảnh đã nén vào sheet '{sheet_name}'.")

                workbook_win32.Save()
                workbook_win32.Close()
                excel.Quit()
            except Exception as e:
                logging.error(f"Lỗi khi thay thế hình ảnh: {str(e)}")
            finally:
                pythoncom.CoUninitialize()
        
        logging.info("Hoàn tất nén ảnh bằng engine Spire.Xls.")
        return True
        
    except Exception as e:
        logging.error(f"Lỗi nghiêm trọng trong quá trình nén ảnh với Spire.Xls: {str(e)}")
        return False
    finally:
        try:
            shutil.rmtree(temp_dir)
            logging.debug("Đã xóa thư mục tạm thời.")
        except:
            logging.warning("Không thể xóa thư mục tạm thời.")
