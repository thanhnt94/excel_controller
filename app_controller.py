# Đường dẫn: excel_toolkit/app_controller.py
# Phiên bản 1.0 - Tách logic xử lý ra khỏi main.py
# Ngày cập nhật: 2025-09-16

import tkinter.filedialog as filedialog
import threading
import logging
import shutil
import tempfile
import os
from excel_controller import ExcelController
from ui import AppUI, TaskSelectionDialog
from ui_notifier import StatusNotifier
from localization import translator
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

class AppController:
    def __init__(self, root):
        self.root = root
        self.ui = AppUI(root, self)
        self.notifier = StatusNotifier(root)
        self.file_paths = []
        
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

    def open_folder(self, folder_path):
        if folder_path and os.path.isdir(folder_path):
            os.startfile(folder_path)

    def open_input_folder(self, event):
        self.open_folder(self.ui.folder_path_entry.get())

    def open_output_folder(self, event):
        self.open_folder(self.ui.output_folder_entry.get())

    def update_main_master_checkbox_state(self):
        if not self.ui.file_checkboxes: return
        all_on = all(cb.get() == 1 for cb in self.ui.file_checkboxes)
        if all_on:
            self.ui.main_master_checkbox.select()
        else:
            self.ui.main_master_checkbox.deselect()

    def toggle_all_files(self):
        new_state_is_on = self.ui.main_master_checkbox_var.get() == "on"
        for checkbox in self.ui.file_checkboxes:
            if new_state_is_on:
                checkbox.select()
            else:
                checkbox.deselect()

    def change_language(self, new_lang_name):
        translator.set_language_by_name(new_lang_name)
        self.ui.update_ui_text()

    def change_log_level(self, new_level_name):
        log_level_map = {
            translator.get_text("log_level_info"): logging.INFO,
            translator.get_text("log_level_debug"): logging.DEBUG
        }
        selected_level = log_level_map.get(new_level_name, logging.INFO)
        logging.getLogger().setLevel(selected_level)
        logging.info(f"Mức độ log đã được thay đổi thành: {new_level_name}")

    def browse_output_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path: 
            self.ui.output_folder_entry.delete(0, "end")
            self.ui.output_folder_entry.insert(0, folder_path)
            
    def log_message(self, message, style='plain', duration=0):
        logging.info(message)
        self.notifier.update_status(message, style=style, duration=duration)

    def browse_folder_event(self):
        folder_path = filedialog.askdirectory()
        if not folder_path: return
        self.ui.folder_path_entry.delete(0, "end")
        self.ui.folder_path_entry.insert(0, folder_path)
        self.log_message(f"Finding files in: {folder_path}", style="process")
        self.file_paths = file_system_ops.get_files_path(folder_path, file_extensions=['.xlsx', '.xlsm', '.xls'], include_subfolders=True)
        if not self.file_paths: 
            self.log_message("No Excel files found.", style="warning")
            return
        self.log_message(f"Found {len(self.file_paths)} files.", style="success")
        self.ui.update_file_list(self.file_paths)

    def run_tasks_event(self):
        selected_files = [self.file_paths[i] for i, cb in enumerate(self.ui.file_checkboxes) if cb.get() == 1]
        if not selected_files: 
            self.log_message("Please select at least one file.", style="warning")
            return
            
        save_mode_text = self.ui.save_option_menu.get()
        save_details = {}
        
        if save_mode_text == translator.get_text("save_rename"):
            affix_type = self.ui.rename_type_var.get()
            affix_text = self.ui.affix_entry.get()
            if not affix_text:
                self.log_message("Vui lòng nhập tiền tố/hậu tố.", style="error")
                return
            save_details['affix_type'] = affix_type
            save_details['affix_text'] = affix_text
        elif save_mode_text == translator.get_text("save_output_folder"):
            folder = self.ui.output_folder_entry.get()
            if not folder or not os.path.isdir(folder): 
                self.log_message("Vui lòng chọn thư mục đích hợp lệ.", style="error")
                return
            save_details['folder'] = folder

        dialog = TaskSelectionDialog(self.root)
        selected_tasks, selected_engine_text, quality_param, selected_label_text = dialog.get_selected_tasks()
        
        engine = None
        if selected_engine_text == translator.get_text("engine_pil"):
            engine = "pil"
        elif selected_engine_text == translator.get_text("engine_spire"):
            engine = "spire"

        if not selected_tasks: 
            self.log_message("Cancelled.", style="info", duration=0)
            return
            
        save_details['text'] = save_mode_text
        self.process_files(selected_files, selected_tasks, engine, quality_param, selected_label_text, save_details)

    def process_files(self, files, tasks, engine, quality_param, label_text, save_details):
        processing_thread = threading.Thread(target=self._run_batch_thread, args=(files, tasks, self.task_map, engine, quality_param, label_text, save_details))
        processing_thread.start()

    def _run_batch_thread(self, files, tasks, task_map, engine, quality_param, label_text, save_details):
        total_files, temp_dir = len(files), tempfile.mkdtemp()
        try:
            self.log_message(f"Processing {total_files} files...", style="process", duration=0)
            for i, original_path in enumerate(files):
                file_name = os.path.basename(original_path)
                temp_path = os.path.join(temp_dir, file_name)
                shutil.copy2(original_path, temp_path)
                
                is_file_processed_successfully = True
                
                with ExcelController(visible=False, optimize_performance=True) as controller:
                    try:
                        if not controller.open_workbook(temp_path):
                            raise Exception(f"Could not open workbook: {file_name}")

                        for task_id in tasks:
                            task_name, task_func = task_map[task_id]
                            self.log_message(f"File {i+1}/{total_files}\nRunning '{task_name}' on: {file_name}", style="info", duration=0)
                            
                            if task_id == "compress_all_images":
                                quality_value = int(quality_param) if quality_param and quality_param.isdigit() else 70
                                task_func(controller, temp_path, engine, quality_value)
                            elif task_id == "add_label":
                                task_func(controller, temp_path, label_text=label_text)
                            else:
                                task_func(controller, temp_path)
                        
                        controller.save_workbook()

                    except Exception as e:
                        self.log_message(f"ERROR processing file: {file_name}\nDetails: {e}", style="error", duration=8)
                        logging.exception(f"An exception occurred while processing {file_name}")
                        is_file_processed_successfully = False
                
                if is_file_processed_successfully:
                    try:
                        mode = save_details['text']
                        if mode == translator.get_text("save_overwrite"):
                            shutil.move(temp_path, original_path)
                            self.log_message(f"Overwrote file: {file_name}", style="success")
                        
                        elif mode == translator.get_text("save_rename"):
                            base, ext = os.path.splitext(original_path)
                            dir_name = os.path.dirname(original_path)
                            affix_text = save_details['affix_text']
                            if save_details['affix_type'] == 'prefix':
                                new_path = os.path.join(dir_name, f"{affix_text}{os.path.basename(base)}{ext}")
                            else: # Suffix
                                new_path = f"{base}{affix_text}{ext}"
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
