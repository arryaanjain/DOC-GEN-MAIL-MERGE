import customtkinter as ctk
from tkinter import messagebox
from tkcalendar import DateEntry
from src.components.status_panel import StatusPanel
from src.backend.app.core.root_dir_setup import get_output_directory
import os
import datetime

from src.utils.theme import AppTheme

class ProcessingPanel(ctk.CTkFrame):
    def __init__(self, master, process_callback):
        super().__init__(
            master,
            **AppTheme.get_frame_style()
        )
        self.process_callback = process_callback
        self.selected_docx = None
        self.output_path = None
        self.root_dir = None

        # Configure grid
        self.grid_columnconfigure(0, weight=1)

        # --- Root Directory Section ---
        self.root_frame = ctk.CTkFrame(self, **AppTheme.get_frame_style())
        self.root_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 4))
        self.root_frame.grid_columnconfigure(1, weight=1)

        self.root_label = ctk.CTkLabel(
            self.root_frame, text="Select Root Directory:",
            **AppTheme.get_label_style(font=AppTheme.FONTS["heading"])
        )
        self.root_label.grid(row=0, column=0, padx=(12, 8), pady=10, sticky="w")
        self.root_entry = ctk.CTkEntry(self.root_frame, **AppTheme.get_input_style())
        self.root_entry.grid(row=0, column=1, padx=(0, 8), pady=10, sticky="ew")
        self.root_button = ctk.CTkButton(
            self.root_frame, text="Browse",
            command=self.select_root_dir,
            fg_color=AppTheme.get_colors()["primary"], 
            hover_color=AppTheme.get_colors()["secondary"], 
            text_color="white",
            **AppTheme.get_button_style()
        )
        self.root_button.grid(row=0, column=2, padx=(0, 12), pady=10, sticky="ew")

        # --- Separator ---
        self.sep1 = ctk.CTkFrame(self, height=2, fg_color=self._get_separator_color())
        self.sep1.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 4))

        # --- Input DOCX Section ---
        self.docx_frame = ctk.CTkFrame(self, **AppTheme.get_frame_style())
        self.docx_frame.grid(row=2, column=0, sticky="ew", padx=12, pady=4)
        self.docx_frame.grid_columnconfigure(1, weight=1)

        self.docx_label = ctk.CTkLabel(
            self.docx_frame, text="Select Input DOCX:",
            **AppTheme.get_label_style(font=AppTheme.FONTS["heading"])
        )
        self.docx_label.grid(row=0, column=0, padx=12, pady=10, sticky="w")
        self.docx_path_label = ctk.CTkLabel(
            self.docx_frame, text="No file selected", 
            **AppTheme.get_label_style()
        )
        self.docx_path_label.grid(row=0, column=1, padx=8, pady=10, sticky="w")
        self.docx_button = ctk.CTkButton(
            self.docx_frame, text="Browse DOCX",
            command=self.select_docx,
            fg_color=AppTheme.get_colors()["primary"], 
            hover_color=AppTheme.get_colors()["secondary"], 
            text_color="white",
            **AppTheme.get_button_style()
        )
        self.docx_button.grid(row=0, column=2, padx=12, pady=10, sticky="ew")

        # --- Separator ---
        self.sep2 = ctk.CTkFrame(self, height=2, fg_color=self._get_separator_color())
        self.sep2.grid(row=3, column=0, sticky="ew", padx=8, pady=(0, 4))

        # --- Processing Date Section ---
        self.date_frame = ctk.CTkFrame(self, **AppTheme.get_frame_style())
        self.date_frame.grid(row=4, column=0, sticky="ew", padx=12, pady=4)
        self.date_frame.grid_columnconfigure(1, weight=1)

        self.date_label = ctk.CTkLabel(
            self.date_frame, text="Processing Date (YYYY-MM-DD):",
            **AppTheme.get_label_style(font=AppTheme.FONTS["heading"])
        )
        self.date_label.grid(row=0, column=0, padx=12, pady=10, sticky="w")
        self.date_entry = DateEntry(self.date_frame, date_pattern='yyyy-mm-dd')
        self.date_entry.grid(row=0, column=1, padx=8, pady=10, sticky="ew")

        # --- Separator ---
        self.sep3 = ctk.CTkFrame(self, height=2, fg_color=self._get_separator_color())
        self.sep3.grid(row=5, column=0, sticky="ew", padx=8, pady=(0, 4))

        # --- Append Excel Section ---
        self.append_frame = ctk.CTkFrame(self, **AppTheme.get_frame_style())
        self.append_frame.grid(row=6, column=0, sticky="ew", padx=12, pady=4)
        self.append_frame.grid_columnconfigure(1, weight=1)

        self.append_label = ctk.CTkLabel(
            self.append_frame, text="Append To Existing Excel (optional):",
            **AppTheme.get_label_style(font=AppTheme.FONTS["heading"])
        )
        self.append_label.grid(row=0, column=0, padx=12, pady=10, sticky="w")
        self.append_path_label = ctk.CTkLabel(
            self.append_frame, text="No file selected", 
            **AppTheme.get_label_style()
        )
        self.append_path_label.grid(row=0, column=1, padx=8, pady=10, sticky="w")
        self.append_button = ctk.CTkButton(
            self.append_frame, text="Browse Excel",
            command=self.select_append_file,
            fg_color=AppTheme.get_colors()["success"], 
            hover_color="#388e3c", 
            text_color="white",
            **AppTheme.get_button_style()
        )
        self.append_button.grid(row=0, column=2, padx=12, pady=10, sticky="ew")
        self.append_file_path = None

        # --- Separator ---
        self.sep4 = ctk.CTkFrame(self, height=2, fg_color=self._get_separator_color())
        self.sep4.grid(row=7, column=0, sticky="ew", padx=8, pady=(0, 4))

        # --- Process Button Section ---
        self.button_frame = ctk.CTkFrame(self, **AppTheme.get_frame_style())
        self.button_frame.grid(row=8, column=0, sticky="ew", padx=12, pady=4)
        self.process_button = ctk.CTkButton(
            self.button_frame, text="Process Document",
            command=self.process_document,
            state="disabled",
            fg_color=AppTheme.get_colors()["warning"], 
            hover_color="#f57c00", 
            text_color="white",
            **AppTheme.get_button_style(font=AppTheme.FONTS["heading"]),
            width=220  # Fixed width for centering
        )
        self.process_button.pack(pady=10, anchor="center")  # Center the button

        # --- Separator ---
        self.sep5 = ctk.CTkFrame(self, height=2, fg_color=self._get_separator_color())
        self.sep5.grid(row=9, column=0, sticky="ew", padx=8, pady=(0, 4))

        # --- Status Panel Section ---
        self.status_panel = StatusPanel(self)
        self.status_panel.grid(row=10, column=0, padx=16, pady=16, sticky="nsew")

    def _get_separator_color(self):
        """Return separator color based on theme."""
        return "#444" if ctk.get_appearance_mode() == "Dark" else "#e0e0e0"

    def select_root_dir(self):
        dir_path = ctk.filedialog.askdirectory()
        if dir_path:
            self.root_dir = dir_path
            self.root_entry.delete(0, "end")
            self.root_entry.insert(0, dir_path)
        self._check_process_ready()

    def select_docx(self):
        file_path = ctk.filedialog.askopenfilename(
            filetypes=[("Word Documents", "*.docx")]
        )
        if file_path:
            self.selected_docx = file_path
            self.docx_path_label.configure(
                text=os.path.basename(file_path),
                text_color=AppTheme.get_colors()["success"]
            )
            self.status_panel.update_status(f"Selected document: {os.path.basename(file_path)}", "info")
        else:
            self.selected_docx = None
            self.docx_path_label.configure(
                text="No file selected",
                text_color=AppTheme.get_colors()["text"]
            )
        self._check_process_ready()

    def select_append_file(self):
        file_path = ctk.filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if file_path:
            self.append_file_path = file_path
            self.append_path_label.configure(
                text=os.path.basename(file_path),
                text_color=AppTheme.get_colors()["success"]
            )
        else:
            self.append_file_path = None
            self.append_path_label.configure(
                text="No file selected",
                text_color=AppTheme.get_colors()["text"]
            )

    def process_document(self):
        if not self.selected_docx or not self.root_dir:
            self.status_panel.update_status("Missing input or output location", "error")
            messagebox.showerror("Error", "Please select both input and output locations")
            return

        current_date = datetime.datetime.now().strftime("%Y-%m-%d")
        output_dir = get_output_directory(self.root_dir, current_date)
        self.output_path = output_dir  # Update output path to the new directory
        
        # Generate timestamp
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_excel = os.path.join(self.output_path, f"converted_{timestamp}.xlsx")
        
        processing_date = self.date_entry.get().strip()  # Get date from entry

        try:
            self.status_panel.update_status("Processing document...", "info")
            self.status_panel.set_progress(0.5)  # Show progress

            # Pass processing_date as third argument
            self.process_callback(self.selected_docx, output_excel, processing_date, self.append_file_path)
            self.status_panel.set_progress(1.0)  # Complete progress
            self.status_panel.update_status("Document processed successfully!", "success")

            messagebox.showinfo("Success", "Document processed successfully!")
        except Exception as e:
            self.status_panel.set_progress(0)  # Reset progress
            self.status_panel.update_status(f"Error: {str(e)}", "error")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def _check_process_ready(self):
        """Enable process button if both input and output are selected"""
        if self.selected_docx and self.root_dir:
            self.process_button.configure(state="normal")
        else:
            self.process_button.configure(state="disabled")