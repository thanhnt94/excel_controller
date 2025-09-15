# Đường dẫn: excel_toolkit/main.py
# Phiên bản 17.0 - Khắc phục tất cả các lỗi và tích hợp đầy đủ tùy chọn nén ảnh
# Ngày cập nhật: 2025-09-15

import customtkinter
import tkinter.filedialog as filedialog
import os
import threading
import logging
import shutil
import tempfile
from datetime import datetime
from excel_controller import ExcelController

from utils import file_system_ops
from processes import (
    set_label, 
    delete_hidden_sheets, 
    delete_external_links, 
    delete_defined_names, 
    set_print_settings,
    clear_excess_cell_formatting,
    compress_all_images,
    refresh_and_clean_pivot_caches
)
from ui_notifier import StatusNotifier
from localization import translator

def setup_logging(level=logging.INFO):
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    log_dir = "logs"
    if not os.path.exists(log_dir): os.makedirs(log_dir)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = os.path.join(log_dir, f"log_{timestamp}.log")
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename, 'w', 'utf-8'),
            logging.StreamHandler()
        ]
    )
    logging.info(f"Hệ thống ghi log đã được khởi tạo. Mức độ: {logging.getLevelName(level)}. File: {log_filename}")

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("blue")
_FILENAME_TRUNCATE_LIMIT = 30

class TaskSelectionDialog(customtkinter.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.transient(parent); self.grab_set()
        self.tasks_vars, self.result = {}, []
        self.engine_var, self.quality_var = None, None
        
        main_frame = customtkinter.CTkFrame(self); main_frame.pack(expand=True, fill="both", padx=20, pady=20)
        
        self.label = customtkinter.CTkLabel(main_frame, font=customtkinter.CTkFont(weight="bold"))
        self.label.pack(pady=(0, 10), anchor="w")
        
        self.tasks_frame = customtkinter.CTkFrame(main_frame, fg_color="transparent")
        self.tasks_frame.pack(fill="both", expand=True)

        self.options_frame = customtkinter.CTkFrame(main_frame, fg_color="transparent")
        self.options_frame.pack(fill="x", pady=10, anchor="w")

        button_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        button_frame.pack(fill="x", padx=20, pady=(15, 20))
        self.ok_button = customtkinter.CTkButton(button_frame, command=self.on_ok)
        self.ok_button.pack(side="right", padx=(10, 0))
        self.cancel_button = customtkinter.CTkButton(button_frame, fg_color="gray", command=self.on_cancel)
        self.cancel_button.pack(side="right")
        self.update_text()
        self.check_options_visibility()

    def check_options_visibility(self, event=None):
        for widget in self.options_frame.winfo_children():
            widget.pack_forget()

        if self.tasks_vars.get("compress_all_images", customtkinter.StringVar()).get() == "on":
            self.engine_label = customtkinter.CTkLabel(self.options_frame, text=translator.get_text("task_compress_all_images_engine_label"), font=customtkinter.CTkFont(weight="bold"))
            self.engine_label.pack(side="left", padx=(10, 5))
            self.engine_var = customtkinter.StringVar(value=translator.get_text("engine_xlwings"))
            self.engine_menu = customtkinter.CTkOptionMenu(self.options_frame, values=[translator.get_text("engine_xlwings"), translator.get_text("engine_spire")], variable=self.engine_var, command=self.update_compression_options)
            self.engine_menu.pack(side="left")

            self.quality_label = customtkinter.CTkLabel(self.options_frame, text=f"DPI:", font=customtkinter.CTkFont(weight="bold"))
            self.quality_label.pack(side="left", padx=(15, 5))
            self.quality_var = customtkinter.StringVar(value="220")
            self.quality_entry = customtkinter.CTkEntry(self.options_frame, width=60, textvariable=self.quality_var)
            self.quality_entry.pack(side="left")

            self.update_compression_options(self.engine_var.get())
        
    def update_compression_options(self, choice):
        if choice == translator.get_text("engine_xlwings"):
            self.quality_label.configure(text="DPI:")
            self.quality_entry.configure(placeholder_text="220")
            self.quality_var.set("220")
        elif choice == translator.get_text("engine_spire"):
            self.quality_label.configure(text=f"{translator.get_text('image_max_size_kb')}:")
            self.quality_entry.configure(placeholder_text="300")
            self.quality_var.set("300")

    def update_text(self):
        self.title(translator.get_text("tasks_dialog_title")); self.geometry("400x400")
        self.label.configure(text=translator.get_text("tasks_dialog_label"))
        
        for widget in self.tasks_frame.winfo_children(): widget.destroy()
        tasks = {
            "add_label": translator.get_text("task_add_label"),
            "delete_hidden_sheets": translator.get_text("task_delete_hidden_sheets"),
            "delete_external_links": translator.get_text("task_delete_external_links"),
            "delete_defined_names": translator.get_text("task_delete_defined_names"),
            "set_print_settings": translator.get_text("task_set_print_settings"),
            "clear_excess_cell_formatting": translator.get_text("task_clear_excess_cell_formatting"),
            "compress_all_images": translator.get_text("task_compress_all_images"),
            "refresh_and_clean_pivot_caches": translator.get_text("task_refresh_and_clean_pivot_caches")
        }
        for task_id, task_name in tasks.items():
            var = customtkinter.StringVar(value=self.tasks_vars.get(task_id, "off"))
            cb = customtkinter.CTkCheckBox(self.tasks_frame, text=task_name, variable=var, onvalue="on", offvalue="off", command=self.check_options_visibility)
            cb.pack(anchor="w", padx=10, pady=5)
            self.tasks_vars[task_id] = var
        self.tasks_vars.get("add_label", customtkinter.StringVar()).set("on")
        self.ok_button.configure(text=translator.get_text("run_button_dialog"))
        self.cancel_button.configure(text=translator.get_text("cancel_button_dialog"))

    def on_ok(self): 
        self.result = [k for k, v in self.tasks_vars.items() if v.get() == "on"]
        if "compress_all_images" in self.result:
            self.engine_var = self.engine_var.get()
            self.quality_var = self.quality_var.get()
        else:
            self.engine_var = None
            self.quality_var = None
        self.destroy()

    def on_cancel(self): 
        self.result = []
        self.destroy()

    def get_selected_tasks(self): 
        self.master.wait_window(self)
        return self.result, self.engine_var, self.quality_var

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("800x550")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) 

        self.main_container = customtkinter.CTkFrame(self, fg_color="transparent")
        self.main_container.grid(row=0, column=0, rowspan=3, padx=20, pady=20, sticky="nsew")
        self.main_container.grid_columnconfigure(0, weight=1)
        self.main_container.grid_rowconfigure(1, weight=1)

        browse_frame = customtkinter.CTkFrame(self.main_container)
        browse_frame.grid(row=0, column=0, sticky="ew")
        browse_frame.grid_columnconfigure(0, weight=1)
        self.folder_path_entry = customtkinter.CTkEntry(browse_frame)
        self.folder_path_entry.grid(row=0, column=0, padx=(10, 5), pady=10, sticky="ew")
        self.browse_button = customtkinter.CTkButton(browse_frame, width=100, command=self.browse_folder_event)
        self.browse_button.grid(row=0, column=1, padx=(5, 10), pady=10)

        files_frame = customtkinter.CTkFrame(self.main_container)
        files_frame.grid(row=1, column=0, pady=(10, 0), sticky="nsew")
        files_frame.grid_columnconfigure(0, weight=1)
        files_frame.grid_rowconfigure(2, weight=1) 

        controls_frame = customtkinter.CTkFrame(files_frame, fg_color="transparent")
        controls_frame.grid(row=0, column=0, padx=10, pady=(5,0), sticky="ew")
        self.select_all_button = customtkinter.CTkButton(controls_frame, width=120, command=self.select_all_files)
        self.select_all_button.pack(side="left", padx=(0, 5))
        self.deselect_all_button = customtkinter.CTkButton(controls_frame, width=120, command=self.deselect_all_files)
        self.deselect_all_button.pack(side="left", padx=(0, 20))
        self.save_label = customtkinter.CTkLabel(controls_frame)
        self.save_label.pack(side="left", padx=(0,5))
        self.save_option_menu = customtkinter.CTkOptionMenu(controls_frame, command=self.update_save_option_widgets)
        self.save_option_menu.pack(side="left")

        self.option_widgets_frame = customtkinter.CTkFrame(files_frame, fg_color="transparent")
        self.option_widgets_frame.grid(row=1, column=0, padx=10, pady=(5,0), sticky="ew")
        self.option_widgets_frame.grid_columnconfigure(0, weight=1)
        self.output_folder_entry = customtkinter.CTkEntry(self.option_widgets_frame)
        self.output_browse_button = customtkinter.CTkButton(self.option_widgets_frame, width=100, command=self.browse_output_folder)
        self.affix_entry = customtkinter.CTkEntry(self.option_widgets_frame)

        self.file_scrollable_frame = customtkinter.CTkScrollableFrame(files_frame)
        self.file_scrollable_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        self.file_scrollable_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.run_button = customtkinter.CTkButton(self.main_container, height=40, font=customtkinter.CTkFont(size=15, weight="bold"), command=self.run_tasks_event)
        self.run_button.grid(row=2, column=0, padx=10, pady=15, sticky="ew")
        
        self.statusbar_frame = customtkinter.CTkFrame(self, height=30, corner_radius=0)
        self.statusbar_frame.grid(row=3, column=0, sticky="ew")
        self.copyright_label = customtkinter.CTkLabel(self.statusbar_frame, text="©KNT15083", font=customtkinter.CTkFont(size=10))
        self.copyright_label.pack(side="right", padx=10, pady=5)
        
        self.lang_menu = customtkinter.CTkOptionMenu(self.statusbar_frame, values=["Tiếng Việt", "English", "日本語"], command=self.change_language)
        self.lang_menu.pack(side="right", padx=10, pady=5)
        self.lang_label = customtkinter.CTkLabel(self.statusbar_frame, font=customtkinter.CTkFont(size=10))
        self.lang_label.pack(side="right", padx=(10, 5), pady=5)
        
        self.log_level_menu = customtkinter.CTkOptionMenu(self.statusbar_frame, command=self.change_log_level)
        self.log_level_menu.pack(side="right", padx=10, pady=5)
        self.log_level_label = customtkinter.CTkLabel(self.statusbar_frame, font=customtkinter.CTkFont(size=10))
        self.log_level_label.pack(side="right", padx=(10, 5), pady=5)

        self.notifier = StatusNotifier(self)
        self.file_paths, self.file_checkboxes = [], []
        
        self.task_map = {
            "add_label": (translator.get_text("task_add_label"), set_label.run),
            "delete_hidden_sheets": (translator.get_text("task_delete_hidden_sheets"), delete_hidden_sheets.run),
            "delete_external_links": (translator.get_text("task_delete_external_links"), delete_external_links.run),
            "delete_defined_names": (translator.get_text("task_delete_defined_names"), delete_defined_names.run),
            "set_print_settings": (translator.get_text("task_set_print_settings"), set_print_settings.run),
            "clear_excess_cell_formatting": (translator.get_text("task_clear_excess_cell_formatting"), clear_excess_cell_formatting.run),
            "compress_all_images": (translator.get_text("task_compress_all_images"), compress_all_images.run),
            "refresh_and_clean_pivot_caches": (translator.get_text("task_refresh_and_clean_pivot_caches"), refresh_and_clean_pivot_caches.run)
        }
        
        self.update_ui_text()

    def change_language(self, new_lang_name):
        translator.set_language_by_name(new_lang_name); self.update_ui_text()

    def change_log_level(self, new_level_name):
        log_level_map = {
            translator.get_text("log_level_info"): logging.INFO,
            translator.get_text("log_level_debug"): logging.DEBUG
        }
        selected_level = log_level_map.get(new_level_name, logging.INFO)
        setup_logging(selected_level)

    def update_ui_text(self):
        self.title(translator.get_text("window_title"))
        self.folder_path_entry.configure(placeholder_text=translator.get_text("browse_placeholder"))
        self.browse_button.configure(text=translator.get_text("browse_button"))
        self.select_all_button.configure(text=translator.get_text("select_all"))
        self.deselect_all_button.configure(text=translator.get_text("deselect_all"))
        self.save_label.configure(text=translator.get_text("save_options_label"))
        
        save_options = [
            translator.get_text("save_overwrite"), translator.get_text("save_backup"),
            translator.get_text("save_prefix"), translator.get_text("save_suffix"), 
            translator.get_text("save_output_folder")
        ]
        current_save_mode = self.save_option_menu.get()
        self.save_option_menu.configure(values=save_options)
        self.save_option_menu.set(save_options[0] if current_save_mode not in save_options else current_save_mode)

        self.output_folder_entry.configure(placeholder_text=translator.get_text("output_folder_placeholder"))
        self.output_browse_button.configure(text=translator.get_text("browse_button"))
        self.affix_entry.configure(placeholder_text=translator.get_text("affix_placeholder"))
        self.file_scrollable_frame.configure(label_text=translator.get_text("file_list_label"))
        self.run_button.configure(text=translator.get_text("run_button"))
        self.lang_label.configure(text=translator.get_text("language_label"))
        
        self.log_level_label.configure(text=translator.get_text("log_level_label"))
        log_levels = [translator.get_text("log_level_info"), translator.get_text("log_level_debug")]
        self.log_level_menu.configure(values=log_levels)
        self.log_level_menu.set(log_levels[0])

        self.update_save_option_widgets()
        
    def update_save_option_widgets(self, choice=None):
        for widget in [self.output_folder_entry, self.output_browse_button, self.affix_entry]:
            widget.grid_remove()
        self.option_widgets_frame.grid_remove()
        
        mode = self.save_option_menu.get()
        if mode == translator.get_text("save_output_folder"):
            self.option_widgets_frame.grid(); self.output_folder_entry.grid(row=0, column=0, sticky="ew")
            self.output_browse_button.grid(row=0, column=1, padx=10)
        elif mode in [translator.get_text("save_prefix"), translator.get_text("save_suffix")]:
            self.option_widgets_frame.grid(); self.affix_entry.grid(row=0, column=0, columnspan=2, sticky="ew")

    def browse_output_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path: self.output_folder_entry.insert(0, folder_path)
            
    def log_message(self, message, style='plain', duration=0):
        logging.info(message); self.notifier.update_status(message, style=style, duration=duration)

    def browse_folder_event(self):
        folder_path = filedialog.askdirectory()
        if not folder_path: return
        self.folder_path_entry.delete(0, "end"); self.folder_path_entry.insert(0, folder_path)
        self.log_message(f"Finding files in: {folder_path}", style="process")
        self.file_paths = file_system_ops.get_files_path(folder_path, file_extensions=['.xlsx', '.xlsm', '.xls'], include_subfolders=True)
        if not self.file_paths: self.log_message("No Excel files found.", style="warning"); return
        self.log_message(f"Found {len(self.file_paths)} files.", style="success")
        self.update_file_list()

    def update_file_list(self):
        self.clear_file_list()
        for i, file_path in enumerate(self.file_paths):
            row, col = divmod(i, 4)
            base_name = os.path.basename(file_path)
            display_name = (base_name[:_FILENAME_TRUNCATE_LIMIT-3] + "...") if len(base_name) > _FILENAME_TRUNCATE_LIMIT else base_name
            checkbox = customtkinter.CTkCheckBox(self.file_scrollable_frame, text=display_name)
            checkbox.grid(row=row, column=col, padx=10, pady=5, sticky="w"); checkbox.select()
            self.file_checkboxes.append(checkbox)

    def clear_file_list(self):
        for widget in self.file_scrollable_frame.winfo_children(): widget.destroy()
        self.file_checkboxes.clear()

    def select_all_files(self):
        for checkbox in self.file_checkboxes: checkbox.select()

    def deselect_all_files(self):
        for checkbox in self.file_checkboxes: checkbox.deselect()

    def run_tasks_event(self):
        selected_files = [self.file_paths[i] for i, cb in enumerate(self.file_checkboxes) if cb.get() == 1]
        if not selected_files: self.log_message("Please select at least one file.", style="warning"); return
            
        save_mode_text, save_details = self.save_option_menu.get(), {}
        
        if save_mode_text == translator.get_text("save_output_folder"):
            folder = self.output_folder_entry.get()
            if not folder or not os.path.isdir(folder): self.log_message("Please select a valid destination folder.", style="error"); return
            save_details['folder'] = folder
        elif save_mode_text in [translator.get_text("save_prefix"), translator.get_text("save_suffix")]:
            affix = self.affix_entry.get()
            if not affix: self.log_message("Please enter a prefix/suffix.", style="error"); return
            save_details['affix'] = affix

        dialog = TaskSelectionDialog(self)
        selected_tasks, selected_engine, quality_param = dialog.get_selected_tasks()
        if not selected_tasks: self.log_message("Cancelled.", style="info", duration=0); return
            
        save_details['text'] = save_mode_text
        self.process_files(selected_files, selected_tasks, selected_engine, quality_param, save_details)

    def process_files(self, files, tasks, engine, quality_param, save_details):
        task_map = {
            "add_label": (translator.get_text("task_add_label"), set_label.run),
            "delete_hidden_sheets": (translator.get_text("task_delete_hidden_sheets"), delete_hidden_sheets.run),
            "delete_external_links": (translator.get_text("task_delete_external_links"), delete_external_links.run),
            "delete_defined_names": (translator.get_text("task_delete_defined_names"), delete_defined_names.run),
            "set_print_settings": (translator.get_text("task_set_print_settings"), set_print_settings.run),
            "clear_excess_cell_formatting": (translator.get_text("task_clear_excess_cell_formatting"), clear_excess_cell_formatting.run),
            "compress_all_images": (translator.get_text("task_compress_all_images"), compress_all_images.run),
            "refresh_and_clean_pivot_caches": (translator.get_text("task_refresh_and_clean_pivot_caches"), refresh_and_clean_pivot_caches.run)
        }
        
        processing_thread = threading.Thread(target=self._run_batch_thread, args=(files, tasks, task_map, engine, quality_param, save_details))
        processing_thread.start()

    def _run_batch_thread(self, files, tasks, task_map, engine, quality_param, save_details):
        total_files, temp_dir = len(files), tempfile.mkdtemp()
        try:
            self.log_message(f"Processing {total_files} files...", style="process", duration=0)
            for i, original_path in enumerate(files):
                file_name = os.path.basename(original_path)
                temp_path = os.path.join(temp_dir, file_name)
                shutil.copy2(original_path, temp_path)
                
                is_file_processed_successfully = True
                for task_id in tasks:
                    task_name, task_func = task_map[task_id]
                    self.log_message(f"File {i+1}/{total_files}\nRunning '{task_name}' on: {file_name}", style="info", duration=0)
                    try:
                        if task_id == "compress_all_images":
                            quality_value = int(quality_param) if quality_param and quality_param.isdigit() else 220
                            task_func(temp_path, engine, quality_value)
                        else:
                            task_func(temp_path)
                    except Exception as e:
                        self.log_message(f"ERROR: {task_name}\nFile: {file_name}\nDetails: {e}", style="error", duration=8)
                        logging.exception(f"An exception occurred while processing {file_name} for task {task_name}")
                        is_file_processed_successfully = False; break 
                
                if is_file_processed_successfully:
                    try:
                        mode = save_details['text']
                        if mode == translator.get_text("save_overwrite"):
                            shutil.move(temp_path, original_path)
                            self.log_message(f"Overwrote file: {file_name}", style="success")
                        elif mode == translator.get_text("save_backup"):
                            base, ext = os.path.splitext(original_path)
                            backup_path = f"{base}_origin{ext}"
                            if os.path.exists(backup_path): os.remove(backup_path)
                            os.rename(original_path, backup_path)
                            shutil.move(temp_path, original_path)
                            self.log_message(f"Processed: {file_name} (original backed up)", style="success")
                        elif mode == translator.get_text("save_prefix"):
                            dir_name = os.path.dirname(original_path)
                            new_path = os.path.join(dir_name, f"{save_details['affix']}{file_name}")
                            shutil.move(temp_path, new_path)
                            self.log_message(f"Saved new file: {os.path.basename(new_path)}", style="success")
                        elif mode == translator.get_text("save_suffix"):
                            base, ext = os.path.splitext(original_path)
                            new_path = f"{base}{save_details['affix']}{ext}"
                            shutil.move(temp_path, new_path)
                            self.log_message(f"Saved new file: {os.path.basename(new_path)}", style="success")
                        elif mode == translator.get_text("save_output_folder"):
                            dest_path = os.path.join(save_details['folder'], file_name)
                            if not os.path.exists(save_details['folder']): os.makedirs(save_details['folder'])
                            shutil.move(temp_path, dest_path)
                            self.log_message(f"Saved to destination: {file_name}", style="success")
                    except Exception as e:
                        self.log_message(f"Error saving file {file_name}: {e}", style="error", duration=8)
                        logging.exception(f"An exception occurred while saving {file_name}")
            self.log_message(f"Completed! Processed {total_files} files.", style="success", duration=5)
        finally:
            shutil.rmtree(temp_dir)

if __name__ == "__main__":
    setup_logging(logging.INFO)
    app = App()
    app.mainloop()
