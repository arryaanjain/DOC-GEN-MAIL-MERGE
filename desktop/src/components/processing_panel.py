import customtkinter as ctk
from tkinter import messagebox
from tkcalendar import Calendar
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

        AppTheme.add_theme_callback(self.update_theme_colors)
        
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
        # Root Directory Button
        self.root_button = ctk.CTkButton(
            self.root_frame, 
            text="Browse",
            command=self.select_root_dir,
            **AppTheme.get_button_style(override_colors=True)
        )
        self.root_button.configure(
            fg_color=AppTheme.get_colors()["primary"],
            hover_color=AppTheme.get_colors()["secondary"]
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
            self.docx_frame, 
            text="Browse DOCX",
            command=self.select_docx,
            **AppTheme.get_button_style(override_colors=True)
        )
        self.docx_button.configure(
            fg_color=AppTheme.get_colors()["primary"],
            hover_color=AppTheme.get_colors()["secondary"]
        )
        self.docx_button.grid(row=0, column=2, padx=12, pady=10, sticky="ew")

        # --- Separator ---
        self.sep2 = ctk.CTkFrame(self, height=2, fg_color=self._get_separator_color())
        self.sep2.grid(row=3, column=0, sticky="ew", padx=8, pady=(0, 4))

                # --- Processing Date Section ---
        self.date_frame = ctk.CTkFrame(self, **AppTheme.get_frame_style())
        self.date_frame.grid(row=4, column=0, sticky="ew", padx=12, pady=4)
        self.date_frame.grid_columnconfigure(1, weight=1)  # Make middle space expand
        
        self.date_label = ctk.CTkLabel(
            self.date_frame, 
            text="Processing Date:",
            **AppTheme.get_label_style(font=AppTheme.FONTS["heading"])
        )
        self.date_label.grid(row=0, column=0, padx=12, pady=10, sticky="w")

        # Date entry using CustomTkinter entry with validation
        self.date_entry = ctk.CTkEntry(
            self.date_frame,
            placeholder_text="YYYY-MM-DD",
            width=120,  # Match button width
            **AppTheme.get_input_style()
        )
        self.date_entry.grid(row=0, column=2, padx=12, pady=10, sticky="e")
        self.date_entry.insert(0, datetime.datetime.now().strftime("%Y-%m-%d"))  # Set current date

        # Add calendar button
        self.calendar_button = ctk.CTkButton(
            self.date_frame,
            text="📅",  # Calendar emoji
            width=40,  # Square button
            command=self.show_calendar,
            **AppTheme.get_button_style(override_colors=True)
        )
        self.calendar_button.configure(
            fg_color=AppTheme.get_colors()["primary"],
            hover_color=AppTheme.get_colors()["secondary"]
        )
        self.calendar_button.grid(row=0, column=3, padx=(0, 12), pady=10)

        # Add validation
        self._validate_date_entry()

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
            self.append_frame, 
            text="Browse Excel",
            command=self.select_append_file,
            **AppTheme.get_button_style(override_colors=True)
        )
        self.append_button.configure(
            fg_color=AppTheme.get_colors()["success"],
            hover_color="#388e3c"
        )
        self.append_button.grid(row=0, column=2, padx=12, pady=10, sticky="ew")
        self.append_file_path = None

        # --- Separator ---
        self.sep4 = ctk.CTkFrame(self, height=2, fg_color=self._get_separator_color())
        self.sep4.grid(row=7, column=0, sticky="ew", padx=8, pady=(0, 4))

        # --- Process Button Section ---
        self.button_frame = ctk.CTkFrame(self, **AppTheme.get_frame_style())
        self.button_frame.grid(row=8, column=0, sticky="ew", padx=12, pady=4)
       
       # Process Button
        self.process_button = ctk.CTkButton(
            self.button_frame, 
            text="Process Document",
            command=self.process_document,
            state="disabled",
            width=220,  # Fixed width for centering
            **AppTheme.get_button_style(
                font=AppTheme.FONTS["heading"],
                override_colors=True
            )
        )
        self.process_button.configure(
            fg_color=AppTheme.get_colors()["warning"],
            hover_color="#f57c00"
        )
        self.process_button.pack(pady=10, anchor="center")  # Center the button

        # --- Separator ---
        self.sep5 = ctk.CTkFrame(self, height=2, fg_color=self._get_separator_color())
        self.sep5.grid(row=9, column=0, sticky="ew", padx=8, pady=(0, 4))

        # --- Status Panel Section ---
        self.status_panel = StatusPanel(self)
        self.status_panel.grid(row=10, column=0, padx=16, pady=16, sticky="nsew")

    def _get_separator_color(self):
        return AppTheme.get_separator_color()

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
    
    def update_theme_colors(self):
        """Update colors when theme changes"""
        colors = AppTheme.get_colors()
        
        # Update buttons
        self.root_button.configure(
            fg_color=colors["primary"],
            hover_color=colors["secondary"]
        )
        self.docx_button.configure(
            fg_color=colors["primary"],
            hover_color=colors["secondary"]
        )
        self.append_button.configure(
            fg_color=colors["success"],
            hover_color="#388e3c"
        )
        self.process_button.configure(
            fg_color=colors["warning"],
            hover_color="#f57c00"
        )
        
        # Update separators
        separator_color = self._get_separator_color()
        self.sep1.configure(fg_color=separator_color)
        self.sep2.configure(fg_color=separator_color)
        self.sep3.configure(fg_color=separator_color)
        self.sep4.configure(fg_color=separator_color)
        self.sep5.configure(fg_color=separator_color)

        # Update labels
        self._update_label_colors()

    def _update_label_colors(self):
        """Update all label colors based on current theme"""
        colors = AppTheme.get_colors()
        
        # Update heading labels
        self.root_label.configure(text_color=colors["text"])
        self.docx_label.configure(text_color=colors["text"])
        self.date_label.configure(text_color=colors["text"])
        self.append_label.configure(text_color=colors["text"])
        
        # Update path labels with appropriate colors based on selection state
        self.docx_path_label.configure(
            text_color=colors["success"] if self.selected_docx else colors["text"]
        )
        self.append_path_label.configure(
            text_color=colors["success"] if self.append_file_path else colors["text"]
        )

    # Then add this method to your ProcessingPanel class
    def show_calendar(self):
        """Show a calendar popup for date selection"""
        def set_date():
            if cal.selection_get():
                selected_date = cal.selection_get().strftime("%Y-%m-%d")
                self.date_entry.delete(0, "end")
                self.date_entry.insert(0, selected_date)
            top.destroy()

        # Create popup window
        top = ctk.CTkToplevel(self)
        top.title("Select Date")
        top.geometry("300x350")
        top.transient(self)  # Make window modal
        top.grab_set()  # Make window modal

        # Create calendar widget
        cal = Calendar(
            top,
            selectmode='day',
            date_pattern='yyyy-mm-dd',
            background=AppTheme.get_colors()["primary"],
            foreground="white",
            selectbackground=AppTheme.get_colors()["secondary"],
            selectforeground="white",
            normalbackground=AppTheme.get_colors()["surface"],
            normalforeground=AppTheme.get_colors()["text"],
            weekendbackground=AppTheme.get_colors()["surface"],
            weekendforeground=AppTheme.get_colors()["text"],
            headersbackground=AppTheme.get_colors()["primary"],
            headersforeground="white"
        )
        cal.pack(expand=True, fill='both', padx=10, pady=10)

        # If there's a current date in the entry, set calendar to that date
        try:
            current_date = datetime.datetime.strptime(
                self.date_entry.get(), "%Y-%m-%d"
            ).date()
            cal.selection_set(current_date)
        except ValueError:
            pass  # Use today's date if current entry is invalid

        # Add select button
        select_btn = ctk.CTkButton(
            top,
            text="Select",
            command=set_date,
            **AppTheme.get_button_style()
        )
        select_btn.pack(pady=10)

    def _validate_date_entry(self):
        """Add validation to date entry"""
        def validate_date(*args):
            date_str = self.date_entry.get()
            try:
                if date_str:
                    datetime.datetime.strptime(date_str, "%Y-%m-%d")
                    self.date_entry.configure(
                        text_color=AppTheme.get_colors()["text"]
                    )
                    return True
            except ValueError:
                self.date_entry.configure(
                    text_color=AppTheme.get_colors()["error"]
                )
                return False

        # Bind validation to entry changes
        self.date_entry.bind('<KeyRelease>', validate_date)
        self.date_entry.bind('<FocusOut>', validate_date)

    def __del__(self):
        """Clean up theme callback when widget is destroyed"""
        AppTheme.remove_theme_callback(self.update_theme_colors)