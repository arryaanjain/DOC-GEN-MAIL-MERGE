import customtkinter as ctk

class AppTheme:
    # Font configurations
    FONTS = {
        "heading": ("Segoe UI", 16, "bold"),
        "normal": ("Segoe UI", 12),
        "small": ("Segoe UI", 10),
        "monospace": ("Consolas", 11)
    }

    @staticmethod
    def get_colors():
        """Return color scheme based on appearance mode"""
        mode = ctk.get_appearance_mode()
        if mode == "Dark":
            return {
                "primary": "#90caf9",
                "secondary": "#1976d2",
                "success": "#81c784",
                "warning": "#ffd54f",
                "error": "#e57373",
                "background": "#23272e",
                "text": "#f5f5f5"
            }
        else:
            return {
                "primary": "#1f538d",
                "secondary": "#14375e",
                "success": "#2ea043",
                "warning": "#d29922",
                "error": "#cb2431",
                "background": "#ffffff",
                "text": "#24292e"
            }

    @staticmethod
    def setup_theme():
        """Configure default theme settings"""
        ctk.set_appearance_mode("system")  # system, light, dark
        ctk.set_default_color_theme("blue")

    @staticmethod
    def get_button_style(font=None):
        """Return common button styling"""
        return {
            "font": font if font else AppTheme.FONTS["normal"],
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
    def get_label_style(font=None):
        """Return common label styling"""
        return {
            "font": font if font else AppTheme.FONTS["normal"],
            "text_color": AppTheme.get_colors()["text"]
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
            "progress_color": AppTheme.get_colors()["primary"]
        }