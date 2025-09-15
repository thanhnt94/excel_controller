# Đường dẫn: ui_notifier.py
# Phiên bản 1.0 - Cập nhật ngày 03/09/2025

import tkinter as tk
from tkinter import font
import queue
import threading
import time
import logging
from dataclasses import dataclass, field
from typing import Dict, Any, Optional, List, Callable, Tuple

# ======================================================================
#       NEW: CONFIGURATION WITH DATACLASSES
# ======================================================================

@dataclass
class NotifierStyle:
    icon: str
    fg: str
    bg: str

@dataclass
class NotifierConfig:
    """Configuration for the StatusNotifier."""
    # --- General ---
    alpha: float = 0.95
    position: str = 'bottom_right'
    margin_x: int = 20
    margin_y: int = 50
    
    # --- Sizing ---
    width: Any = 'auto'
    height: Any = 'auto'
    min_width: int = 300
    max_width: int = 450
    min_height: int = 70
    
    # --- Font & Text ---
    font_family: str = 'Segoe UI'
    font_size: int = 10
    font_style: str = 'normal'
    font_color: str = 'auto'
    
    # --- Layout & Icons ---
    padding_x: int = 20
    padding_y: int = 15
    icon_text_spacing: int = 10
    show_icons: bool = True
    
    # --- Border configuration ---
    border_thickness: int = 1
    border_color: str = '#FFFFFF'
    
    # --- Behavior ---
    default_duration: int = 0
    default_style: str = 'info'
    
    # --- Animation ---
    animation: str = 'fade'
    animation_speed: int = 10
    
    # --- Style Definitions ---
    styles: Dict[str, NotifierStyle] = field(default_factory=lambda: {
        'plain':    NotifierStyle(icon='',    fg='#FFFFFF', bg='#34495E'),
        'info':     NotifierStyle(icon='ℹ️',    fg='#E1F5FE', bg='#0288D1'),
        'success':  NotifierStyle(icon='✅',    fg='#FFFFFF', bg='#27AE60'),
        'warning':  NotifierStyle(icon='⚠️',    fg='#000000', bg='#F39C12'),
        'error':    NotifierStyle(icon='❌',    fg='#FFFFFF', bg='#C0392B'),
        'process':  NotifierStyle(icon='⚙️',    fg='#FFFFFF', bg='#7F8C8D'),
        'question': NotifierStyle(icon='❓',    fg='#FFFFFF', bg='#8E44AD'),
        'debug':    NotifierStyle(icon='🐞',    fg='#AAB7B8', bg='#17202A'),
        'download': NotifierStyle(icon='📥',    fg='#FFFFFF', bg='#16A085'),
        'upload':   NotifierStyle(icon='📤',    fg='#FFFFFF', bg='#16A085'),
        'auth':     NotifierStyle(icon='🔑',    fg='#FFFFFF', bg='#D35400'),
    })

def _update_dataclass_from_dict(dc_instance, user_dict):
    """Helper to merge a dict into a dataclass instance."""
    for key, value in user_dict.items():
        if hasattr(dc_instance, key):
            if key == 'styles' and isinstance(value, dict):
                for style_name, style_dict in value.items():
                    if style_name in dc_instance.styles and isinstance(style_dict, dict):
                        for sk, sv in style_dict.items():
                            if hasattr(dc_instance.styles[style_name], sk):
                                setattr(dc_instance.styles[style_name], sk, sv)
                    elif isinstance(style_dict, dict):
                          dc_instance.styles[style_name] = NotifierStyle(**style_dict)
            else:
                setattr(dc_instance, key, value)
    return dc_instance


class StatusNotifier:
    """
    Quản lý một cửa sổ thông báo không chặn, sử dụng cấu hình dataclass.
    """
    def __init__(self, parent_root: tk.Tk, config: Optional[Dict[str, Any]] = None):
        """
        Khởi tạo một instance StatusNotifier.

        Tham số:
            parent_root (tk.Tk): Cửa sổ gốc của ứng dụng, nơi tất cả các widget khác được đặt.
            config (Dict, tùy chọn): Từ điển cấu hình tùy chỉnh.
        """
        self.parent_root = parent_root
        self.queue = queue.Queue()
        
        base_config = NotifierConfig()
        self.config = _update_dataclass_from_dict(base_config, config or {})
        
        self.root: Optional[tk.Tk] = None
        self._hide_job: Optional[str] = None
        self._animation_job: Optional[str] = None
        
        self._is_paused: bool = False
        self._start_time: float = 0
        self._current_duration: float = 0
        
        self._buttons: List[tk.Button] = []
        
        self._setup_gui()
        self.parent_root.after(50, self._check_queue)

    def _setup_gui(self):
        """Khởi tạo các widget giao diện."""
        self.root = tk.Toplevel(self.parent_root)
        self.root.overrideredirect(True)
        self.root.wm_attributes("-topmost", True)
        self.root.wm_attributes("-alpha", 0)  # Bắt đầu ẩn hoàn toàn
        self.root.withdraw()
        
        font_style_str = self.config.font_style.lower()
        weight = 'bold' if 'bold' in font_style_str else 'normal'
        slant = 'italic' if 'italic' in font_style_str else 'roman'

        self.icon_font = font.Font(family=self.config.font_family, size=self.config.font_size + 4, weight='bold')
        self.text_font = font.Font(family=self.config.font_family, size=self.config.font_size, weight=weight, slant=slant)
        self.button_font = font.Font(family=self.config.font_family, size=self.config.font_size -1, weight='bold')

        self.border_frame = tk.Frame(self.root, bg=self.config.border_color, bd=0)
        self.border_frame.pack(expand=True, fill='both')

        self.main_frame = tk.Frame(self.border_frame, bd=0)
        self.main_frame.pack(expand=True, fill='both', padx=self.config.border_thickness, pady=self.config.border_thickness)

        self.content_frame = tk.Frame(self.main_frame)
        self.content_frame.pack(side='top', fill='x', expand=True)
        
        self.buttons_frame = tk.Frame(self.main_frame)
        self.buttons_frame.pack(side='bottom', fill='x', pady=(5,0))

        self.icon_label = tk.Label(self.content_frame, font=self.icon_font, justify='center')
        self.text_label = tk.Label(self.content_frame, font=self.text_font, justify='left')
        
        widgets_to_bind = [self.border_frame, self.main_frame, self.content_frame, self.icon_label, self.text_label]
        for widget in widgets_to_bind:
            widget.bind("<Button-1>", self._dismiss)
            widget.bind("<Enter>", self._on_mouse_enter)
            widget.bind("<Leave>", self._on_mouse_leave)

    def _check_queue(self):
        """Kiểm tra hàng đợi để xử lý các tác vụ thông báo."""
        try:
            while True:
                task = self.queue.get_nowait()
                if self._hide_job: self.root.after_cancel(self._hide_job); self._hide_job = None
                if self._animation_job: self.root.after_cancel(self._animation_job); self._animation_job = None
                if task['command'] == "STOP": self._animate_out(self.config.animation, destroy_after=True); break
                elif task['command'] == "UPDATE": self._process_update(task['data'])
        except queue.Empty:
            pass
        if self.root and self.root.winfo_exists(): self.parent_root.after(50, self._check_queue)

    def _process_update(self, data: Dict[str, Any]):
        """Cập nhật nội dung và hiển thị thông báo."""
        style_config = self.config.styles.get(data['style'], self.config.styles['info'])
        bg_color = style_config.bg
        fg_color = self.config.font_color if self.config.font_color != 'auto' else style_config.fg
        
        self.border_frame.config(bg=self.config.border_color)
        self.main_frame.config(bg=bg_color)
        self.content_frame.config(bg=bg_color)
        self.buttons_frame.config(bg=bg_color)

        self.text_label.config(text=data['text'], bg=bg_color, fg=fg_color)
        self.icon_label.pack_forget()
        self.text_label.pack_forget()

        icon_text = style_config.icon if self.config.show_icons else ''
        if icon_text:
            self.icon_label.config(text=icon_text, bg=bg_color, fg=fg_color)
            self.icon_label.pack(side='left', fill='y', padx=(self.config.padding_x, self.config.icon_text_spacing), pady=self.config.padding_y)
        
        self.text_label.pack(side='left', fill='both', expand=True, padx=(0 if icon_text else self.config.padding_x, self.config.padding_x), pady=self.config.padding_y)

        for button in self._buttons: button.destroy()
        self._buttons.clear()

        buttons_data = data.get('buttons')
        if buttons_data:
            self.buttons_frame.pack(side='bottom', fill='x', padx=self.config.padding_x, pady=(0, self.config.padding_y))
            for button_info in buttons_data:
                btn = tk.Button(
                    self.buttons_frame, text=button_info['text'], font=self.button_font,
                    bg=fg_color, fg=bg_color, relief='flat', overrelief='raised',
                    borderwidth=1, command=lambda cmd=button_info['command']: self._on_button_click(cmd)
                )
                btn.pack(side='right', padx=(5, 0))
                self._buttons.append(btn)
        else:
            self.buttons_frame.pack_forget()

        self.root.update_idletasks()
        
        icon_width = self.icon_label.winfo_reqwidth() if icon_text else 0
        wraplength = self.config.max_width - (self.config.padding_x * 2) - self.config.icon_text_spacing - icon_width - (self.config.border_thickness * 2)
        self.text_label.config(wraplength=wraplength)
        self.root.update_idletasks()
        
        req_width = self.main_frame.winfo_reqwidth()
        req_height = self.main_frame.winfo_reqheight()

        final_width = int(max(self.config.min_width, min(req_width, self.config.max_width)))
        final_height = int(max(self.config.min_height, min(req_height, self.parent_root.winfo_screenheight())))
        
        animation = data.get('animation') or self.config.animation
        self._animate_in(final_width, final_height, animation)

        duration = data['duration']
        if duration > 0:
            self._is_paused = False
            self._current_duration = duration
            self._start_time = time.time()
            self._hide_job = self.root.after(int(duration * 1000), lambda: self._animate_out(animation))

    def _on_mouse_enter(self, event=None):
        if self._hide_job:
            self._is_paused = True
            self.root.after_cancel(self._hide_job)
            self._hide_job = None
            elapsed_time = time.time() - self._start_time
            self._current_duration -= elapsed_time

    def _on_mouse_leave(self, event=None):
        if self._is_paused:
            self._is_paused = False
            if self._current_duration > 0:
                self._start_time = time.time()
                animation = self.config.animation
                self._hide_job = self.root.after(int(self._current_duration * 1000), lambda: self._animate_out(animation))
    
    def _on_button_click(self, command: Optional[Callable]):
        if command:
            try: command()
            except Exception as e: logging.error(f"Error executing button command: {e}", exc_info=True)
        self._dismiss()

    def _dismiss(self, event=None):
        if self._hide_job: self.root.after_cancel(self._hide_job); self._hide_job = None
        if self._animation_job: self.root.after_cancel(self._animation_job); self._animation_job = None
        self._animate_out(self.config.animation)

    def _get_positions(self, width: int, height: int, animation_style: str) -> Tuple[int, int, int, int]:
        screen_width = self.parent_root.winfo_screenwidth()
        screen_height = self.parent_root.winfo_screenheight()
        margin_x, margin_y = self.config.margin_x, self.config.margin_y
        pos_map = {
            'top_right': (screen_width - width - margin_x, margin_y),
            'top_left': (margin_x, margin_y),
            'bottom_right': (screen_width - width - margin_x, screen_height - height - margin_y),
            'bottom_left': (margin_x, screen_height - height - margin_y),
            'center': ((screen_width // 2) - (width // 2), (screen_height // 2) - (height // 2))
        }
        end_x, end_y = pos_map.get(self.config.position, pos_map['bottom_right'])
        start_x, start_y = end_x, end_y
        if 'slide' in animation_style:
            if 'up' in animation_style: start_y = screen_height
            elif 'down' in animation_style: start_y = -height
            elif 'left' in animation_style: start_x = screen_width
            elif 'right' in animation_style: start_x = -width
        return start_x, start_y, end_x, end_y

    def _animate_in(self, width: int, height: int, animation: str):
        self.root.deiconify() # Hiển thị cửa sổ trước khi bắt đầu animation
        start_x, start_y, end_x, end_y = self._get_positions(width, height, animation)
        self.root.geometry(f'{width}x{height}+{start_x}+{start_y}')
        
        if animation == 'none':
            self.root.attributes("-alpha", self.config.alpha)
            self.root.geometry(f'{width}x{height}+{end_x}+{end_y}')
            return
        
        total_steps = 20
        def step(i):
            progress = i / total_steps
            new_x = int(start_x + (end_x - start_x) * progress)
            new_y = int(start_y + (end_y - start_y) * progress)
            
            if 'fade' in animation:
                self.root.attributes("-alpha", self.config.alpha * progress)
            
            if 'grow' in animation:
                scale = progress
                current_w, current_h = int(width * scale), int(height * scale)
                pos_x, pos_y = end_x + (width - current_w) // 2, end_y + (height - current_h) // 2
                self.root.geometry(f'{current_w}x{current_h}+{pos_x}+{pos_y}')
                if 'fade' not in animation:
                    self.root.attributes("-alpha", self.config.alpha * progress)
            else:
                self.root.geometry(f'+{new_x}+{new_y}')
            
            if i >= total_steps:
                self.root.geometry(f'{width}x{height}+{end_x}+{end_y}')
                self.root.attributes("-alpha", self.config.alpha)
                self._animation_job = None
            else:
                self._animation_job = self.root.after(self.config.animation_speed, lambda: step(i + 1))
        
        step(1)

    def _animate_out(self, animation: str, destroy_after: bool = False):
        width, height = self.root.winfo_width(), self.root.winfo_height()
        current_x, current_y = self.root.winfo_x(), self.root.winfo_y()
        start_x, start_y, target_x, target_y = self._get_positions(width, height, animation)
        
        if animation == 'none':
            self.root.withdraw()
            if destroy_after: self.root.destroy()
            return

        total_steps = 20
        def step(i):
            progress = i / total_steps
            new_x = int(current_x + (target_x - current_x) * progress)
            new_y = int(current_y + (target_y - current_y) * progress)
            
            if 'fade' in animation or 'grow' in animation:
                self.root.attributes("-alpha", self.config.alpha * (1 - progress))
            
            if 'grow' in animation:
                scale = 1 - progress
                current_w, current_h = int(width * scale), int(height * scale)
                pos_x, pos_y = current_x + (width - current_w) // 2, current_y + (height - current_h) // 2
                self.root.geometry(f'{current_w}x{current_h}+{pos_x}+{pos_y}')
            else:
                self.root.geometry(f'+{new_x}+{new_y}')
            
            if i >= total_steps:
                self.root.withdraw()
                self._animation_job = None
                if destroy_after: self.root.destroy()
            else:
                self._animation_job = self.root.after(self.config.animation_speed, lambda: step(i + 1))
        
        step(1)

    def update_status(self, text: str, style: Optional[str] = None, duration: Optional[int] = None, animation: Optional[str] = None, buttons: Optional[List[Dict[str, Any]]] = None):
        """
        Gửi một thông báo mới vào hàng đợi.

        Tham số:
            text (str): Nội dung của thông báo.
            style (str, tùy chọn): Kiểu thông báo ('info', 'success', 'warning', 'error', 'process', v.v.).
            duration (int, tùy chọn): Thời gian hiển thị thông báo (giây). 0 để hiển thị vĩnh viễn cho đến khi nhấp.
            animation (str, tùy chọn): Kiểu animation ('fade', 'slide_up', 'none').
            buttons (List[Dict], tùy chọn): Danh sách các nút hành động, mỗi nút là một dictionary.
        """
        if duration is None: duration = self.config.default_duration
        if style is None: style = self.config.default_style
        task_data = {'text': text, 'style': style, 'duration': duration, 'animation': animation, 'buttons': buttons}
        self.queue.put({'command': 'UPDATE', 'data': task_data})

    def stop(self):
        """Dừng tất cả các hoạt động của notifier và đóng cửa sổ."""
        self.queue.put({'command': 'STOP'})


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    
    print("--- Running Notifier Demonstration (v13.1) ---")
    
    root = tk.Tk()
    root.title("Main App Window")
    root.geometry("200x200")

    def retry_operation():
        print(">>> ACTION: 'Thử lại' button was clicked. Running retry logic...")
        notifier.update_status("Đang thử lại kết nối...", style='process', duration=3)

    notifier = StatusNotifier(root, {'position': 'bottom_right', 'animation': 'fade'})
    
    def run_demo():
        print("\n1. Testing Pause on Hover and Click to Dismiss...")
        notifier.update_status("Di chuột vào để tạm dừng. Nhấp chuột để đóng.", style='question', duration=8)
        
        root.after(9000, lambda: print("\n2. Testing Action Buttons..."))
        root.after(9000, lambda: print("    - Một thông báo lỗi với các nút hành động sẽ xuất hiện."))
        root.after(9000, lambda: print("    - Hãy thử nhấn nút 'Thử lại' hoặc 'Hủy'."))

        action_buttons = [
            {'text': 'Thử lại', 'command': retry_operation},
            {'text': 'Hủy', 'command': None}
        ]
        
        root.after(10000, lambda: notifier.update_status(
            text="Không thể lưu tệp. Vui lòng kiểm tra lại quyền truy cập.",
            style='error',
            duration=0,
            buttons=action_buttons
        ))
        
        root.after(20000, notifier.stop)

    run_demo()
    root.mainloop()
