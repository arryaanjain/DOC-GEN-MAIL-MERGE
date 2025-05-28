import customtkinter as ctk
from tkinter import messagebox
from src.components.status_panel import StatusPanel
import os

class ProcessingPanel(ctk.CTkFrame):
    def __init__(self, master, process_callback):
        super().__init__(master)
        self.process_callback = process_callback
        self.selected_docx = None
        self.output_path = None
        
        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        
        # File selection section
        self.docx_label = ctk.CTkLabel(self, text="Select Input DOCX:")
        self.docx_label.grid(row=0, column=0, padx=10, pady=(10,0), sticky="w")
        
        self.docx_button = ctk.CTkButton(
            self, 
            text="Browse DOCX",
            command=self.select_docx
        )
        self.docx_button.grid(row=1, column=0, columnspan=2, padx=10, pady=(5,10), sticky="ew")
        
        self.docx_path_label = ctk.CTkLabel(self, text="No file selected", text_color="gray")
        self.docx_path_label.grid(row=2, column=0, columnspan=2, padx=10, pady=(0,10), sticky="w")
        
        # Output selection section
        self.output_label = ctk.CTkLabel(self, text="Select Output Location:")
        self.output_label.grid(row=3, column=0, padx=10, pady=(10,0), sticky="w")
        
        self.output_button = ctk.CTkButton(
            self, 
            text="Choose Output Location",
            command=self.select_output
        )
        self.output_button.grid(row=4, column=0, columnspan=2, padx=10, pady=(5,10), sticky="ew")
        
        self.output_path_label = ctk.CTkLabel(self, text="No location selected", text_color="gray")
        self.output_path_label.grid(row=5, column=0, columnspan=2, padx=10, pady=(0,10), sticky="w")
        
        # Processing button
        self.process_button = ctk.CTkButton(
            self,
            text="Process Document",
            command=self.process_document,
            state="disabled"
        )
        self.process_button.grid(row=6, column=0, columnspan=2, padx=10, pady=20, sticky="ew")
        self.status_panel = StatusPanel(self)
        self.status_panel.grid(row=7, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
    
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

    def select_output(self):
        output_dir = ctk.filedialog.askdirectory()
        if output_dir:
            self.output_path = output_dir
            self.output_path_label.configure(
                text=os.path.basename(output_dir),
                text_color="green"
            )
            self.status_panel.update_status(f"Selected output directory: {os.path.basename(output_dir)}", "info")
            self._check_process_ready()

    def process_document(self):
        if not self.selected_docx or not self.output_path:
            self.status_panel.update_status("Missing input or output location", "error")
            messagebox.showerror("Error", "Please select both input and output locations")
            return
            
        # Generate output Excel filename based on input docx
        base_name = os.path.splitext(os.path.basename(self.selected_docx))[0]
        output_excel = os.path.join(self.output_path, f"{base_name}.xlsx")
        
        try:
            self.status_panel.update_status("Processing document...", "info")
            self.status_panel.set_progress(0.5)  # Show progress
            
            self.process_callback(self.selected_docx, output_excel)
            self.status_panel.set_progress(1.0)  # Complete progress
            self.status_panel.update_status("Document processed successfully!", "success")
            
            messagebox.showinfo("Success", "Document processed successfully!")
        except Exception as e:
            self.status_panel.set_progress(0)  # Reset progress
            self.status_panel.update_status(f"Error: {str(e)}", "error")
            
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def _check_process_ready(self):
        """Enable process button if both input and output are selected"""
        if self.selected_docx and self.output_path:
            self.process_button.configure(state="normal")
        else:
            self.process_button.configure(state="disabled")