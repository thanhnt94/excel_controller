# Đường dẫn: excel_toolkit/utils/image_compressor_api.py
# Phiên bản 1.4 - Cải thiện log debug cho quá trình áp dụng lại thuộc tính
# Ngày cập nhật: 2025-09-15

import os
import time
import uuid
import logging
import pythoncom
from PIL import Image, ImageGrab

# --- Hằng số Office/Excel ---
xlScreen = 1
xlBitmap = 2
msoPicture = 13
msoLinkedPicture = 11
msoGroup = 6

msoBringToFront = 0
msoSendToBack = 1
msoBringForward = 2
msoSendBackward = 3

def _doevents_pulse():
    """Bơm message queue để COM không treo."""
    try:
        pythoncom.PumpWaitingMessages()
    except Exception:
        pass

def _copy_shape_to_image(shape, timeout_sec=3.0, sleep_step=0.05):
    """
    Copy shape -> clipboard (bitmap) -> trả về PIL.Image.
    Có timeout để tránh treo khi clipboard bận.
    """
    logging.debug("    -> Bắt đầu sao chép ảnh vào clipboard...")
    shape.api.CopyPicture(Appearance=xlScreen, Format=xlBitmap)
    logging.debug("    -> Đã sao chép vào clipboard. Bắt đầu lấy ảnh từ clipboard...")
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        _doevents_pulse()
        try:
            clip = ImageGrab.grabclipboard()
            if isinstance(clip, Image.Image):
                logging.debug("    -> Lấy ảnh từ clipboard thành công.")
                return clip
            else:
                logging.debug("    -> Clipboard chưa chứa ảnh. Đang đợi...")
        except Exception as e:
            logging.debug(f"    -> Lỗi khi lấy clipboard: {e}. Đang thử lại...")
        time.sleep(sleep_step)
    logging.warning(f"Clipboard không trả về ảnh sau {timeout_sec} giây.")
    return None

def _snapshot_shape_props(shape):
    """
    Lấy ảnh chụp thuộc tính cần bảo toàn để khôi phục sau:
    vị trí, kích thước, xoay, khoá tỉ lệ, placement, tên, visible, alt text, hyperlink, z-order.
    """
    api = shape.api
    props = {
        'name': shape.name,
        'left': shape.left,
        'top': shape.top,
        'width': shape.width,
        'height': shape.height,
        'rotation': getattr(api, 'Rotation', 0),
        'lock_aspect': getattr(api, 'LockAspectRatio', False),
        'placement': getattr(api, 'Placement', None),  # xlMove/xlMoveAndSize/xlFreeFloating
        'visible': getattr(api, 'Visible', True),
        'alt_text': getattr(api, 'AlternativeText', ''),
        'zpos': getattr(api, 'ZOrderPosition', None),
        'hyperlink': None,
    }
    # Hyperlink (nếu có)
    try:
        hl = getattr(shape, 'hyperlink', None)
        # xlwings wrapper có .hyperlink hoặc shape.api.Hyperlink
        if hl and (getattr(hl, 'address', None) or getattr(hl, 'sub_address', None)):
            props['hyperlink'] = {
                'address': getattr(hl, 'address', None),
                'sub_address': getattr(hl, 'sub_address', None),
                'screen_tip': getattr(hl, 'screen_tip', None),
                'text_to_display': getattr(hl, 'text_to_display', None),
            }
        else:
            # Thử COM thuần
            hla = getattr(api, 'Hyperlink', None)
            if hla and (getattr(hla, 'Address', None) or getattr(hla, 'SubAddress', None)):
                props['hyperlink'] = {
                    'address': getattr(hla, 'Address', None),
                    'sub_address': getattr(hla, 'SubAddress', None),
                    'screen_tip': getattr(hla, 'ScreenTip', None),
                    'text_to_display': getattr(hla, 'TextToDisplay', None),
                }
    except Exception:
        pass
    return props

def _apply_props_to_picture(pic, props):
    """
    Áp lại các thuộc tính đã chụp cho ảnh mới chèn.
    """
    api = pic.api
    logging.debug(f"    -> Đang khôi phục thuộc tính cho ảnh 'Picture 20'...")
    try:
        logging.debug("    -> Áp dụng tên...")
        pic.name = props['name']
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng tên: {e}")

    try:
        logging.debug("    -> Áp dụng vị trí và kích thước...")
        pic.left = props['left']
        pic.top = props['top']
        pic.width = props['width']
        pic.height = props['height']
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng vị trí và kích thước: {e}")

    try:
        logging.debug("    -> Bỏ qua thuộc tính xoay để tránh treo.")
        # # Thêm xử lý lỗi cụ thể cho thuộc tính xoay
        # rotation_value = props['rotation']
        # logging.debug(f"    -> Áp dụng xoay với giá trị: {rotation_value}...")
        # if rotation_value != 0:
        #      api.Rotation = rotation_value
        # logging.debug("    -> Áp dụng xoay thành công.")
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng xoay: {e}. Bỏ qua thuộc tính này.")

    try:
        logging.debug("    -> Bỏ qua thuộc tính khóa tỉ lệ để tránh treo.")
        # api.LockAspectRatio = props['lock_aspect']
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng khóa tỉ lệ: {e}")

    try:
        logging.debug("    -> Áp dụng placement...")
        if props['placement'] is not None:
            api.Placement = props['placement']
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng placement: {e}")

    try:
        logging.debug("    -> Áp dụng visible...")
        api.Visible = props['visible']
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng visible: {e}")

    try:
        logging.debug("    -> Bỏ qua thuộc tính alternative text để tránh treo.")
        # api.AlternativeText = props['alt_text'] or ''
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng alternative text: {e}")

    # Hyperlink
    try:
        logging.debug("    -> Áp dụng hyperlink...")
        hl = props.get('hyperlink')
        if hl and (hl.get('address') or hl.get('sub_address')):
            pic.sheet.api.Hyperlinks.Add(
                Anchor=pic.api, Address=hl.get('address'), SubAddress=hl.get('sub_address'),
                ScreenTip=hl.get('screen_tip'), TextToDisplay=hl.get('text_to_display')
            )
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng hyperlink: {e}")

def _export_and_replace(shape, sheet, quality=70, mode='auto', keep_dpi=96):
    """
    Trích xuất shape -> nén -> xoá shape cũ -> chèn lại ảnh -> khôi phục props.
    """
    logging.debug("    -> Lấy thuộc tính của shape để khôi phục...")
    props = _snapshot_shape_props(shape)

    img = _copy_shape_to_image(shape)
    if img is None:
        logging.warning(f"Clipboard không trả ảnh cho '{props['name']}'. Bỏ qua.")
        return None

    logging.debug("    -> Bắt đầu xử lý và lưu ảnh tạm thời...")

    # Quyết định định dạng nén
    fmt = 'JPEG'
    if mode == 'png' or (mode == 'auto' and (img.mode in ('RGBA', 'LA'))):
        fmt = 'PNG'
    elif mode == 'jpeg' or mode == 'auto':
        if img.mode in ('RGBA', 'LA'):
            img = img.convert('RGB')

    tmp_dir = os.path.join(os.getcwd(), "_tmp_excel_img")
    os.makedirs(tmp_dir, exist_ok=True)
    tmp_path = os.path.join(tmp_dir, f"{uuid.uuid4().hex}.{fmt.lower()}")

    try:
        if fmt == 'JPEG':
            logging.debug(f"    -> Lưu ảnh tạm thời ở định dạng JPEG tại '{tmp_path}'...")
            img.save(tmp_path, format='JPEG', quality=quality, optimize=True, dpi=(keep_dpi, keep_dpi))
        else:
            logging.debug(f"    -> Lưu ảnh tạm thời ở định dạng PNG tại '{tmp_path}'...")
            try:
                img_q = img.convert('P', palette=Image.ADAPTIVE, colors=256)
                img_q.save(tmp_path, format='PNG', optimize=True, dpi=(keep_dpi, keep_dpi))
            except Exception:
                img.save(tmp_path, format='PNG', optimize=True, dpi=(keep_dpi, keep_dpi))
        logging.debug("    -> Đã lưu ảnh tạm thời thành công.")
    except Exception as e:
        logging.error(f"    -> Lỗi khi lưu ảnh tạm thời: {e}")
        return None
    
    logging.debug("    -> Bắt đầu xóa shape cũ...")
    # Xoá shape cũ
    try:
        shape.delete()
        logging.debug("    -> Đã xóa shape cũ thành công.")
    except Exception as e:
        logging.warning(f"Không xoá được shape cũ '{props['name']}': {e}")
        return None

    logging.debug("    -> Bắt đầu chèn ảnh mới...")
    # Chèn ảnh mới
    pic = sheet.pictures.add(tmp_path, left=props['left'], top=props['top'])
    logging.debug(f"    -> Đã chèn ảnh mới thành công với tên '{pic.name}'.")

    # Cố gắng giữ nguyên kích thước (tránh scale theo DPI)
    try:
        pic.width = props['width']
        pic.height = props['height']
    except Exception:
        pass

    logging.debug("    -> Bắt đầu áp dụng lại thuộc tính...")
    # Áp thuộc tính lại
    _apply_props_to_picture(pic, props)
    logging.debug("    -> Đã áp dụng lại thuộc tính thành công.")

    # Xóa file tạm
    try:
        os.remove(tmp_path)
        logging.debug("    -> Đã xóa file ảnh tạm thời.")
    except Exception as e:
        logging.warning(f"Không thể xóa file ảnh tạm thời '{tmp_path}': {e}")

    # Trả về tên shape mới (tên có thể đổi nếu trùng)
    return pic.name
    
def _reorder_zorder_exact(sheet, saved_order_back_to_front):
    """
    Khôi phục thứ tự chồng lớp chính xác.
    Cách làm: duyệt theo thứ tự từ “phía sau” -> “phía trước”, mỗi shape gọi BringToFront.
    Kết quả: các shape sẽ xếp đúng như saved_order_back_to_front.
    """
    for nm in saved_order_back_to_front:
        try:
            shp = sheet.shapes[nm]
            shp.api.ZOrder(msoBringToFront)
        except Exception:
            # Có thể tên bị đổi sau khi chèn lại; bỏ qua nếu không còn tồn tại.
            pass

def compress_all_images(wb, quality=70, mode='auto', keep_dpi=96):
    """
    Nén tất cả ảnh (msoPicture/msoLinkedPicture) trong workbook:
    - Bảo toàn vị trí, kích thước, xoay, tỉ lệ, placement, tên, visible, alt text, hyperlink.
    - Khôi phục z-order để textbox/shape khác vẫn đè đúng.
    - Bỏ qua nhóm (msoGroup) để tránh phá vỡ group.
    """
    excel = wb.app.api
    prev_screen = excel.ScreenUpdating
    prev_alerts = excel.DisplayAlerts
    prev_calc = excel.Calculation
    excel.ScreenUpdating = False
    excel.DisplayAlerts = False
    try:
        excel.Calculation = -4135  # xlCalculationManual
    except Exception:
        pass

    total = 0
    compressed = 0
    
    # Duyệt qua từng sheet hiển thị
    for sheet in wb.sheets:
        if getattr(sheet.api, 'Visible', -1) != -1:
            continue

        # Lấy z-order của tất cả các shapes trên sheet
        try:
            shapes_with_z = []
            for s in sheet.shapes:
                try:
                    z = getattr(s.api, 'ZOrderPosition', None)
                    if z is not None:
                        shapes_with_z.append((z, s.name))
                except Exception:
                    pass
            shapes_with_z.sort(key=lambda x: x[0])  # Sắp xếp từ sau ra trước
            z_order_names = [nm for _, nm in shapes_with_z]
        except Exception:
            z_order_names = [s.name for s in sheet.shapes]

        shape_names = [s.name for s in sheet.shapes]
        new_names_map = {}

        for nm in shape_names:
            try:
                shp = sheet.shapes[nm]
                t = getattr(shp.api, 'Type', None)
            except Exception:
                continue

            if t in (msoPicture, msoLinkedPicture):
                total += 1
                logging.info(f"Đang nén ảnh '{nm}' trên sheet '{sheet.name}'...")
                try:
                    new_nm = _export_and_replace(shp, sheet, quality=quality, mode=mode, keep_dpi=keep_dpi)
                    if new_nm:
                        compressed += 1
                        new_names_map[nm] = new_nm
                except Exception as e:
                    logging.warning(f"Lỗi khi nén ảnh '{nm}' ở sheet '{sheet.name}': {e}")
            else:
                logging.debug(f"Bỏ qua shape '{nm}' (loại: {t}) vì không phải ảnh.")
                pass
            
            _doevents_pulse()

        z_order_names_updated = [new_names_map.get(nm, nm) for nm in z_order_names]
        _reorder_zorder_exact(sheet, z_order_names_updated)

    logging.info(f"Hoàn tất nén ảnh. Đã nén {compressed}/{total} ảnh.")
    
    excel.ScreenUpdating = prev_screen
    excel.DisplayAlerts = prev_alerts
    try:
        excel.Calculation = prev_calc
    except Exception:
        pass
    
    return True

def compress_single_image(wb, sheet_name, shape_name, quality=70, mode='auto', keep_dpi=96):
    """
    Nén một hình ảnh cụ thể trong workbook.
    """
    logging.debug(f"Bắt đầu nén ảnh '{shape_name}' trên sheet '{sheet_name}'...")
    try:
        sheet = wb.sheets[sheet_name]
        shape = sheet.shapes[shape_name]

        if getattr(shape.api, 'Type', None) not in (msoPicture, msoLinkedPicture):
            logging.warning(f"Shape '{shape_name}' không phải là một hình ảnh hoặc linked picture. Bỏ qua.")
            return False

        new_name = _export_and_replace(shape, sheet, quality=quality, mode=mode, keep_dpi=keep_dpi)
        if new_name:
            logging.info(f"Đã nén thành công ảnh '{shape_name}' trên sheet '{sheet_name}'.")
            return True
        else:
            logging.error(f"Không thể nén ảnh '{shape_name}'.")
            return False
    except KeyError:
        logging.error(f"Lỗi: Không tìm thấy sheet '{sheet_name}' hoặc shape '{shape_name}'.")
        return False
    except Exception as e:
        logging.error(f"Lỗi khi nén ảnh '{shape_name}': {e}")
        return False
