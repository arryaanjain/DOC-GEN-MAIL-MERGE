import customtkinter as ctk

class AppTheme:
    # Color schemes
    COLORS = {
        "primary": "#1f538d",
        "secondary": "#14375e",
        "success": "#2ea043",
        "warning": "#d29922",
        "error": "#cb2431",
        "background": "#ffffff",
        "text": "#24292e"
    }
    
    # Font configurations
    FONTS = {
        "heading": ("Segoe UI", 16, "bold"),
        "normal": ("Segoe UI", 12),
        "small": ("Segoe UI", 10),
        "monospace": ("Consolas", 11)
    }
    
    @staticmethod
    def setup_theme():
        """Configure default theme settings"""
        ctk.set_appearance_mode("system")  # system, light, dark
        ctk.set_default_color_theme("blue")
    
    @staticmethod
    def get_button_style():
        """Return common button styling"""
        return {
            "font": AppTheme.FONTS["normal"],
            "corner_radius": 6,
            "border_width": 0,
            "hover": True
        }
    
    @staticmethod
    def get_frame_style():
        """Return common frame styling"""
        return {
            "corner_radius": 10,
            "border_width": 1,
            "fg_color": "transparent"
        }
    
    @staticmethod
    def get_label_style():
        """Return common label styling"""
        return {
            "font": AppTheme.FONTS["normal"],
            "text_color": AppTheme.COLORS["text"]
        }
    
    @staticmethod
    def get_input_style():
        """Return common input styling"""
        return {
            "font": AppTheme.FONTS["normal"],
            "corner_radius": 6,
            "border_width": 1
        }
    
    @staticmethod
    def get_progress_style():
        """Return progress bar styling"""
        return {
            "corner_radius": 6,
            "border_width": 0,
            "progress_color": AppTheme.COLORS["primary"]
        }