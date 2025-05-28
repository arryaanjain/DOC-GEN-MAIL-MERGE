import customtkinter as ctk
from datetime import datetime

class StatusPanel(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        
        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        
        # Status label
        self.status_label = ctk.CTkLabel(
            self,
            text="Ready",
            text_color="gray",
            font=("Segoe UI", 12)
        )
        self.status_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        
        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.grid(row=1, column=0, padx=10, pady=(0, 5), sticky="ew")
        self.progress_bar.set(0)
        
        # Log frame
        self.log_frame = ctk.CTkTextbox(
            self,
            height=100,
            font=("Consolas", 11),
            wrap="word"
        )
        self.log_frame.grid(row=2, column=0, padx=10, pady=5, sticky="nsew")

    def update_status(self, message: str, status_type: str = "info"):
        """Update status with different types: info, success, error, warning"""
        color_map = {
            "info": "gray",
            "success": "green",
            "error": "red",
            "warning": "orange"
        }
        self.status_label.configure(
            text=message,
            text_color=color_map.get(status_type, "gray")
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