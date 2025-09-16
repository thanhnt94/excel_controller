# Đường dẫn: excel_toolkit/ui.py
# Phiên bản 1.3 - Sửa lỗi khoảng trống layout trong giao diện chính
# Ngày cập nhật: 2025-09-16

import customtkinter
import tkinter as tk
import os
from localization import translator

_FILENAME_TRUNCATE_LIMIT = 30

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event):
        if self.tooltip_window or not self.text:
            return
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tooltip_window, text=self.text, justify='left',
                      background="#333333", foreground="white", relief='solid', borderwidth=1,
                      font=("Segoe UI", 10, "normal"))
        label.pack(ipadx=5, ipady=3)

    def hide_tooltip(self, event):
        if self.tooltip_window:
            self.tooltip_window.destroy()
        self.tooltip_window = None

class TaskSelectionDialog(customtkinter.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.transient(parent); self.grab_set()
        self.tasks_vars, self.result = {}, []
        self.engine_var, self.quality_var, self.label_text_var = None, None, None
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        main_frame = customtkinter.CTkFrame(self)
        main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)
        
        self.label = customtkinter.CTkLabel(main_frame, font=customtkinter.CTkFont(weight="bold"))
        self.label.grid(row=0, column=0, pady=(0, 10), sticky="w")

        self.master_checkbox_var = customtkinter.StringVar(value="off")
        self.master_checkbox = customtkinter.CTkCheckBox(main_frame, command=self.toggle_all_tasks, variable=self.master_checkbox_var, onvalue="on", offvalue="off")
        self.master_checkbox.grid(row=1, column=0, padx=10, pady=(0, 5), sticky="w")
        
        self.tasks_container = customtkinter.CTkFrame(main_frame, fg_color="transparent")
        self.tasks_container.grid(row=2, column=0, sticky="nsew")
        self.tasks_container.grid_columnconfigure(0, weight=1)

        self.options_frame = customtkinter.CTkFrame(main_frame, fg_color="transparent")
        self.options_frame.grid(row=3, column=0, pady=(5,0), sticky="ew")

        button_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="ew")
        button_frame.grid_columnconfigure(0, weight=1)
        self.cancel_button = customtkinter.CTkButton(button_frame, fg_color="gray", command=self.on_cancel)
        self.cancel_button.pack(side="right")
        self.ok_button = customtkinter.CTkButton(button_frame, command=self.on_ok)
        self.ok_button.pack(side="right", padx=(0, 10))
        
        self.update_text()
        self.check_options_visibility()

    def update_master_checkbox_state(self):
        all_tasks = self.tasks_vars.values()
        if not all_tasks: return
        all_on = all(var.get() == "on" for var in all_tasks)
        
        if all_on:
            self.master_checkbox.select()
        else:
            self.master_checkbox.deselect()

    def toggle_all_tasks(self):
        new_state = self.master_checkbox_var.get()
        for var in self.tasks_vars.values():
            var.set(new_state)
        self.check_options_visibility()

    def on_task_changed(self):
        self.update_master_checkbox_state()
        self.check_options_visibility()
        
    def check_options_visibility(self, event=None):
        for widget in self.options_frame.winfo_children():
            widget.destroy()

        if self.tasks_vars.get("add_label", customtkinter.StringVar()).get() == "on":
            label_frame = customtkinter.CTkFrame(self.options_frame, fg_color="transparent")
            label_frame.pack(fill="x", pady=(0, 5), anchor="w")

            label_option_label = customtkinter.CTkLabel(label_frame, text=translator.get_text("task_add_label_option_label"), font=customtkinter.CTkFont(weight="bold"))
            label_option_label.pack(side="left", padx=(10, 5))
            
            label_options = ["Nissan Confidential C", "Nissan Confidential A", "Nissan Confidential B"]
            self.label_text_var = customtkinter.StringVar(value=label_options[0])
            label_option_menu = customtkinter.CTkOptionMenu(label_frame, values=label_options, variable=self.label_text_var)
            label_option_menu.pack(side="left")

        if self.tasks_vars.get("compress_all_images", customtkinter.StringVar()).get() == "on":
            compress_frame = customtkinter.CTkFrame(self.options_frame, fg_color="transparent")
            compress_frame.pack(fill="x", pady=(5, 0), anchor="w")
            
            compress_option_label = customtkinter.CTkLabel(compress_frame, text=translator.get_text("task_compress_all_images_engine_label"), font=customtkinter.CTkFont(weight="bold"))
            compress_option_label.pack(side="left", padx=(10, 5))
            
            self.engine_var = customtkinter.StringVar(value=translator.get_text("engine_pil"))
            self.engine_menu = customtkinter.CTkOptionMenu(compress_frame, values=[translator.get_text("engine_pil"), translator.get_text("engine_spire")], variable=self.engine_var, command=self.update_compression_options)
            self.engine_menu.pack(side="left")

            self.quality_label = customtkinter.CTkLabel(compress_frame, text="", font=customtkinter.CTkFont(weight="bold"))
            self.quality_label.pack(side="left", padx=(15, 5))
            self.quality_var = customtkinter.StringVar()
            self.quality_entry = customtkinter.CTkEntry(compress_frame, width=60, textvariable=self.quality_var)
            self.quality_entry.pack(side="left")

            self.update_compression_options(self.engine_var.get())
        
    def update_compression_options(self, choice):
        if choice == translator.get_text("engine_pil"):
            self.quality_label.configure(text="Chất lượng (1-95):")
            self.quality_entry.configure(placeholder_text="70")
            self.quality_var.set("70")
        elif choice == translator.get_text("engine_spire"):
            self.quality_label.configure(text=f"{translator.get_text('image_max_size_kb')}:")
            self.quality_entry.configure(placeholder_text="300")
            self.quality_var.set("300")

    def update_text(self):
        self.title(translator.get_text("tasks_dialog_title"))
        self.label.configure(text=translator.get_text("tasks_dialog_label"))
        self.master_checkbox.configure(text=translator.get_text("select_deselect_all"))
        
        tasks_structure = {
            "category_cleanup": {
                "delete_hidden_sheets": translator.get_text("task_delete_hidden_sheets"),
                "delete_external_links": translator.get_text("task_delete_external_links"),
                "delete_defined_names": translator.get_text("task_delete_defined_names"),
            },
            "category_optimization": {
                "clear_excess_cell_formatting": translator.get_text("task_clear_excess_cell_formatting"),
                "compress_all_images": translator.get_text("task_compress_all_images"),
                "refresh_and_clean_pivot_caches": translator.get_text("task_refresh_and_clean_pivot_caches"),
            },
            "category_utilities": {
                "add_label": translator.get_text("task_add_label"),
                "set_print_settings": translator.get_text("task_set_print_settings"),
            }
        }
        
        for widget in self.tasks_container.winfo_children(): widget.destroy()
        self.tasks_vars = {} 

        for cat_key, tasks in tasks_structure.items():
            category_name = translator.get_text(cat_key)
            
            category_label = customtkinter.CTkLabel(self.tasks_container, text=category_name, font=customtkinter.CTkFont(weight="bold"), anchor="w")
            category_label.pack(fill="x", padx=5, pady=(8, 0))

            for task_id, task_name in tasks.items():
                var = customtkinter.StringVar(value="off")
                task_frame = customtkinter.CTkFrame(self.tasks_container, fg_color="transparent")
                task_frame.pack(fill="x", padx=(20,0))
                cb = customtkinter.CTkCheckBox(task_frame, text=task_name, variable=var, onvalue="on", offvalue="off", command=self.on_task_changed)
                cb.pack(anchor="w", padx=10, pady=2)
                self.tasks_vars[task_id] = var
        
        self.update_master_checkbox_state()
        self.ok_button.configure(text=translator.get_text("run_button_dialog"))
        self.cancel_button.configure(text=translator.get_text("cancel_button_dialog"))
        
        self.update_idletasks()
        self.geometry(f"570x{self.winfo_reqheight()}")
        self.resizable(False, False)

    def on_ok(self): 
        self.result = [k for k, v in self.tasks_vars.items() if v.get() == "on"]
        
        if "compress_all_images" in self.result:
            self.engine_var = self.engine_menu.get()
            self.quality_var = self.quality_var.get()
        else:
            self.engine_var = None
            self.quality_var = None

        if "add_label" in self.result and self.label_text_var:
            self.label_text_var = self.label_text_var.get()
        else:
            self.label_text_var = None

        self.destroy()

    def on_cancel(self): 
        self.result = []
        self.destroy()

    def get_selected_tasks(self): 
        self.master.wait_window(self)
        return self.result, self.engine_var, self.quality_var, self.label_text_var

class AppUI:
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.file_checkboxes = []

        self.root.geometry("550x550")
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

        browse_frame = customtkinter.CTkFrame(self.root)
        browse_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        browse_frame.grid_columnconfigure(0, weight=1)
        
        self.input_folder_label = customtkinter.CTkLabel(browse_frame, text="", cursor="hand2")
        self.input_folder_label.grid(row=0, column=0, columnspan=2, padx=10, pady=(5,0), sticky="w")
        self.input_folder_label.bind("<Button-1>", self.controller.open_input_folder)
        self.input_folder_label.bind("<Enter>", lambda e: self.on_folder_label_enter(e, self.input_folder_label))
        self.input_folder_label.bind("<Leave>", lambda e: self.on_folder_label_leave(e, self.input_folder_label, "input_folder_label"))

        self.folder_path_entry = customtkinter.CTkEntry(browse_frame)
        self.folder_path_entry.grid(row=1, column=0, padx=(10, 5), pady=10, sticky="ew")
        self.browse_button = customtkinter.CTkButton(browse_frame, width=100, command=self.controller.browse_folder_event)
        self.browse_button.grid(row=1, column=1, padx=(5, 10), pady=10)

        files_frame = customtkinter.CTkFrame(self.root)
        files_frame.grid(row=1, column=0, padx=20, pady=0, sticky="nsew")
        files_frame.grid_columnconfigure(0, weight=1)
        files_frame.grid_rowconfigure(2, weight=1) 

        controls_frame = customtkinter.CTkFrame(files_frame, fg_color="transparent")
        controls_frame.grid(row=0, column=0, padx=10, pady=(5,0), sticky="ew")
        
        self.main_master_checkbox_var = customtkinter.StringVar(value="on")
        self.main_master_checkbox = customtkinter.CTkCheckBox(controls_frame, command=self.controller.toggle_all_files, variable=self.main_master_checkbox_var, onvalue="on", offvalue="off")
        self.main_master_checkbox.pack(side="left")

        save_options_frame = customtkinter.CTkFrame(controls_frame, fg_color="transparent")
        save_options_frame.pack(side="right")

        self.save_label = customtkinter.CTkLabel(save_options_frame)
        self.save_label.pack(side="left", padx=(0,5))
        self.save_option_menu = customtkinter.CTkOptionMenu(save_options_frame, command=self.update_save_option_widgets)
        self.save_option_menu.pack(side="left")

        self.option_widgets_frame = customtkinter.CTkFrame(files_frame, fg_color="transparent")
        # Khung này sẽ được quản lý (grid/grid_remove) trong update_save_option_widgets

        self.file_scrollable_frame = customtkinter.CTkScrollableFrame(files_frame)
        self.file_scrollable_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        self.file_scrollable_frame.grid_columnconfigure((0, 1), weight=1)

        self.run_button = customtkinter.CTkButton(self.root, height=40, font=customtkinter.CTkFont(size=15, weight="bold"), command=self.controller.run_tasks_event, fg_color="#27AE60", hover_color="#2ECC71")
        self.run_button.grid(row=2, column=0, padx=20, pady=(15,20), sticky="ew")
        
        self.statusbar_frame = customtkinter.CTkFrame(self.root, height=30, corner_radius=0)
        self.statusbar_frame.grid(row=3, column=0, sticky="ew")
        self.statusbar_frame.grid_columnconfigure(1, weight=1) 

        statusbar_left_frame = customtkinter.CTkFrame(self.statusbar_frame, fg_color="transparent")
        statusbar_left_frame.grid(row=0, column=0, sticky="w", padx=10)
        self.lang_label = customtkinter.CTkLabel(statusbar_left_frame, font=customtkinter.CTkFont(size=10))
        self.lang_label.pack(side="left", padx=(0, 5), pady=5)
        self.lang_menu = customtkinter.CTkOptionMenu(statusbar_left_frame, values=["Tiếng Việt", "English", "日本語"], command=self.controller.change_language)
        self.lang_menu.pack(side="left", pady=5)
        self.log_level_label = customtkinter.CTkLabel(statusbar_left_frame, font=customtkinter.CTkFont(size=10))
        self.log_level_label.pack(side="left", padx=(10, 5), pady=5)
        self.log_level_menu = customtkinter.CTkOptionMenu(statusbar_left_frame, command=self.controller.change_log_level)
        self.log_level_menu.pack(side="left", pady=5)

        self.copyright_label = customtkinter.CTkLabel(self.statusbar_frame, text="©KNT15083", font=customtkinter.CTkFont(size=10))
        self.copyright_label.grid(row=0, column=2, sticky="e", padx=10, pady=5)
        
        self.default_label_color = customtkinter.ThemeManager.theme["CTkLabel"]["text_color"]
        self.update_ui_text()
    
    def on_folder_label_enter(self, event, label):
        label.configure(text=translator.get_text("open_folder_hover_label"), text_color="#6495ED")

    def on_folder_label_leave(self, event, label, text_key):
        label.configure(text=translator.get_text(text_key), text_color=self.default_label_color)

    def update_ui_text(self):
        self.root.title(translator.get_text("window_title"))
        self.input_folder_label.configure(text=translator.get_text("input_folder_label"))
        self.folder_path_entry.configure(placeholder_text=translator.get_text("browse_placeholder"))
        self.browse_button.configure(text=translator.get_text("browse_button"))
        self.main_master_checkbox.configure(text=translator.get_text("select_deselect_all"))
        self.save_label.configure(text=translator.get_text("save_options_label"))
        
        save_options = [
            translator.get_text("save_overwrite"),
            translator.get_text("save_rename"),
            translator.get_text("save_output_folder")
        ]
        current_save_mode = self.save_option_menu.get()
        self.save_option_menu.configure(values=save_options)
        self.save_option_menu.set(save_options[0] if current_save_mode not in save_options else current_save_mode)
        
        self.file_scrollable_frame.configure(label_text=translator.get_text("file_list_label"))
        self.run_button.configure(text=translator.get_text("run_button"))
        self.lang_label.configure(text=translator.get_text("language_label"))
        
        self.log_level_label.configure(text=translator.get_text("log_level_label"))
        log_levels = [translator.get_text("log_level_info"), translator.get_text("log_level_debug")]
        self.log_level_menu.configure(values=log_levels)
        self.log_level_menu.set(log_levels[0])

        self.update_save_option_widgets()
        
    def update_save_option_widgets(self, choice=None):
        # SỬA LỖI: Luôn ẩn frame trước
        self.option_widgets_frame.grid_remove()
        for widget in self.option_widgets_frame.winfo_children():
            widget.destroy()

        mode = self.save_option_menu.get()
        if mode in [translator.get_text("save_rename"), translator.get_text("save_output_folder")]:
            # Chỉ hiển thị lại frame khi cần
            self.option_widgets_frame.grid(row=1, column=0, padx=10, pady=(5,0), sticky="ew")

        if mode == translator.get_text("save_rename"):
            self.rename_type_var = customtkinter.StringVar(value="prefix")
            
            prefix_radio = customtkinter.CTkRadioButton(self.option_widgets_frame, text=translator.get_text("save_rename_prefix"), variable=self.rename_type_var, value="prefix")
            prefix_radio.grid(row=0, column=0, padx=10, pady=5, sticky="w")
            
            suffix_radio = customtkinter.CTkRadioButton(self.option_widgets_frame, text=translator.get_text("save_rename_suffix"), variable=self.rename_type_var, value="suffix")
            suffix_radio.grid(row=0, column=1, padx=10, pady=5, sticky="w")
            
            self.affix_entry = customtkinter.CTkEntry(self.option_widgets_frame, placeholder_text=translator.get_text("affix_placeholder"))
            self.affix_entry.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

        elif mode == translator.get_text("save_output_folder"):
            output_folder_label = customtkinter.CTkLabel(self.option_widgets_frame, text=translator.get_text("output_folder_label"), cursor="hand2")
            output_folder_label.grid(row=0, column=0, columnspan=2, padx=10, pady=(5,0), sticky="w")
            output_folder_label.bind("<Button-1>", self.controller.open_output_folder)
            output_folder_label.bind("<Enter>", lambda e: self.on_folder_label_enter(e, output_folder_label))
            output_folder_label.bind("<Leave>", lambda e: self.on_folder_label_leave(e, output_folder_label, "output_folder_label"))
            
            self.output_folder_entry = customtkinter.CTkEntry(self.option_widgets_frame, placeholder_text=translator.get_text("output_folder_placeholder"))
            self.output_folder_entry.grid(row=1, column=0, columnspan=2, padx=(10,5), pady=10, sticky="ew")
            
            self.output_browse_button = customtkinter.CTkButton(self.option_widgets_frame, width=100, text=translator.get_text("browse_button"), command=self.controller.browse_output_folder)
            self.output_browse_button.grid(row=1, column=2, padx=(5,10), pady=10)

    def update_file_list(self, file_paths):
        self.clear_file_list()
        for i, file_path in enumerate(file_paths):
            row, col = divmod(i, 2)
            base_name = os.path.basename(file_path)
            display_name = (base_name[:_FILENAME_TRUNCATE_LIMIT-3] + "...") if len(base_name) > _FILENAME_TRUNCATE_LIMIT else base_name
            
            cell_frame = customtkinter.CTkFrame(self.file_scrollable_frame, fg_color="transparent")
            cell_frame.grid(row=row, column=col, padx=5, pady=2, sticky="ew")

            checkbox = customtkinter.CTkCheckBox(cell_frame, text=display_name, command=self.controller.update_main_master_checkbox_state)
            checkbox.pack(side="left", padx=(5,0))
            
            self.file_checkboxes.append(checkbox)
            ToolTip(checkbox, text=file_path)
            
        self.main_master_checkbox.select()
        self.controller.toggle_all_files()
        self.controller.update_main_master_checkbox_state()

    def clear_file_list(self):
        for widget in self.file_scrollable_frame.winfo_children(): widget.destroy()
        self.file_checkboxes.clear()

