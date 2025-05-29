import customtkinter as ctk
from tkinter import messagebox
from tkcalendar import DateEntry
from src.components.status_panel import StatusPanel
from src.backend.app.core.root_dir_setup import get_output_directory
import os
import datetime

class ProcessingPanel(ctk.CTkFrame):
    def __init__(self, master, process_callback):
        super().__init__(master)
        self.process_callback = process_callback
        self.selected_docx = None
        self.output_path = None
        self.root_dir = None

        # Configure grid
        self.grid_columnconfigure(0, weight=0)  # Label
        self.grid_columnconfigure(1, weight=1)  # Entry expands
        self.grid_columnconfigure(2, weight=0)  # Button

        # Root directory selection section
        self.root_label = ctk.CTkLabel(self, text="Select Root Directory:")
        self.root_label.grid(row=0, column=0, padx=(10, 5), pady=(10, 0), sticky="w")
        self.root_entry = ctk.CTkEntry(self)
        self.root_entry.grid(row=0, column=1, padx=(0, 5), pady=(10, 0), sticky="ew")
        self.root_button = ctk.CTkButton(
            self,
            text="Browse",
            command=self.select_root_dir
        )
        self.root_button.grid(row=0, column=2, padx=(0, 10), pady=(10, 0), sticky="ew")
        # File selection section
        self.docx_label = ctk.CTkLabel(self, text="Select Input DOCX:")
        self.docx_label.grid(row=1, column=0, padx=10, pady=(10,0), sticky="w")
        self.docx_button = ctk.CTkButton(
            self, 
            text="Browse DOCX",
            command=self.select_docx
        )
        self.docx_button.grid(row=1, column=2, padx=10, pady=(10,0), sticky="ew")
        self.docx_path_label = ctk.CTkLabel(self, text="No file selected", text_color="gray")
        self.docx_path_label.grid(row=1, column=1, columnspan=2, padx=10, pady=(10,0), sticky="w")

        # Processing date section
        self.date_label = ctk.CTkLabel(self, text="Processing Date (YYYY-MM-DD):")
        self.date_label.grid(row=7, column=0, padx=10, pady=(0,0), sticky="w")
        self.date_entry = DateEntry(self, date_pattern='yyyy-mm-dd')
        self.date_entry.grid(row=7, column=1, padx=10, pady=(0,0), sticky="ew")

        # Append-to file selection section (moved above status panel)
        self.append_label = ctk.CTkLabel(self, text="Append To Existing Excel (optional):")
        self.append_label.grid(row=8, column=0, padx=10, pady=(10,0), sticky="w")
        self.append_button = ctk.CTkButton(
            self,
            text="Browse Excel",
            command=self.select_append_file
        )
        self.append_button.grid(row=8, column=2, padx=10, pady=(10,0), sticky="ew")
        self.append_path_label = ctk.CTkLabel(self, text="No file selected", text_color="gray")
        self.append_path_label.grid(row=8, column=1, columnspan=2, padx=10, pady=(10,0), sticky="w")
        self.append_file_path = None

        # Processing button
        self.process_button = ctk.CTkButton(
            self,
            text="Process Document",
            command=self.process_document,
            state="disabled"
        )
        self.process_button.grid(row=11, column=1, padx=10, pady=20, sticky="ew")

        # Status panel (now at the bottom)
        self.status_panel = StatusPanel(self)
        self.status_panel.grid(row=12, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

    def select_root_dir(self):
        dir_path = ctk.filedialog.askdirectory()
        if dir_path:
            self.root_dir = dir_path
            self.root_entry.delete(0, "end")
            self.root_entry.insert(0, dir_path)

    def select_docx(self):
        file_path = ctk.filedialog.askopenfilename(
            filetypes=[("Word Documents", "*.docx")]
        )
        if file_path:
            self.selected_docx = file_path
            self.docx_path_label.configure(
                text=os.path.basename(file_path),
                text_color="green"
            )
            self.status_panel.update_status(f"Selected document: {os.path.basename(file_path)}", "info")
            self._check_process_ready()

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

    def select_append_file(self):
            file_path = ctk.filedialog.askopenfilename(
                filetypes=[("Excel Files", "*.xlsx")]
            )
            if file_path:
                self.append_file_path = file_path
                self.append_path_label.configure(
                    text=os.path.basename(file_path),
                    text_color="green"
                )
            else:
                self.append_file_path = None
                self.append_path_label.configure(
                    text="No file selected",
                    text_color="gray"
                )