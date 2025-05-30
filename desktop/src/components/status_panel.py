import customtkinter as ctk
from datetime import datetime
from src.utils.theme import AppTheme


class StatusPanel(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, **AppTheme.get_frame_style())
        
        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        
        # Status label
        self.status_label = ctk.CTkLabel(
            self,
            text="Ready",
            **AppTheme.get_label_style(is_secondary=True)
        )
        self.status_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        
        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(
            self,
            **AppTheme.get_progress_style()
        )
        self.progress_bar.grid(row=1, column=0, padx=10, pady=(0, 5), sticky="ew")
        self.progress_bar.set(0)
        
        # Log frame
        self.log_frame = ctk.CTkTextbox(
            self,
            height=100,
            **AppTheme.get_textbox_style()
        )
        self.log_frame.grid(row=2, column=0, padx=10, pady=5, sticky="nsew")

    def _log_message(self, message: str, level: str):
        """Add message to log with timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {level.upper()}: {message}\n"
        
        # Get theme colors for log levels
        colors = AppTheme.get_colors()
        tag_colors = {
            "info": colors["text_secondary"],
            "success": colors["success"],
            "error": colors["error"],
            "warning": colors["warning"]
        }
        
        # Insert log entry
        self.log_frame.configure(state="normal")
        self.log_frame.insert("end", log_entry)
        self.log_frame.configure(state="disabled")
        self.log_frame.see("end")  # Auto-scroll to bottom

    def update_status(self, message: str, status_type: str = "info"):
        """Update status with different types: info, success, error, warning"""
        colors = AppTheme.get_colors()
        color_map = {
            "info": colors["text_secondary"],
            "success": colors["success"],
            "error": colors["error"],
            "warning": colors["warning"]
        }
        self.status_label.configure(
            text=message,
            text_color=color_map.get(status_type, colors["text_secondary"])
        )
        self._log_message(message, status_type)

    def set_progress(self, value: float):
        """Update progress bar (value between 0 and 1)"""
        self.progress_bar.set(value)

    def _log_message(self, message: str, level: str):
        """Add message to log with timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {level.upper()}: {message}\n"
        
        # Configure tag colors
        tag_colors = {
            "info": "gray",
            "success": "green",
            "error": "red",
            "warning": "orange"
        }
        
        # Insert log entry
        self.log_frame.configure(state="normal")
        self.log_frame.insert("end", log_entry)
        self.log_frame.configure(state="disabled")
        self.log_frame.see("end")  # Auto-scroll to bottom

    def clear_log(self):
        """Clear the log textbox"""
        self.log_frame.configure(state="normal")
        self.log_frame.delete("1.0", "end")
        self.log_frame.configure(state="disabled")