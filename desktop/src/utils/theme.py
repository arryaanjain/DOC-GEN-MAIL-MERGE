import customtkinter as ctk

class AppTheme:
    # Add class variable to store callbacks
    _theme_callbacks = []
    _root = None  # Store root window reference
    # Font configurations
    FONTS = {
        "heading": ("Segoe UI", 16, "bold"),
        "normal": ("Segoe UI", 12),
        "small": ("Segoe UI", 10),
        "monospace": ("Consolas", 11)
    }

    # Color schemes for both modes
    COLORS = {
        "Dark": {
            "primary": "#90caf9",      # Light blue
            "secondary": "#1976d2",     # Darker blue
            "success": "#81c784",       # Light green
            "warning": "#ffd54f",       # Light yellow
            "error": "#e57373",         # Light red
            "background": "#23272e",    # Dark background
            "surface": "#2d323b",       # Slightly lighter than background
            "text": "#f5f5f5",          # Almost white
            "text_secondary": "#a0a0a0", # Lighter gray
            "border": "#404040",        # Dark gray
            "separator": "#444444"      # Slightly lighter than border
        },
        "Light": {
            "primary": "#1f538d",       # Dark blue
            "secondary": "#14375e",      # Darker blue
            "success": "#2ea043",       # Green
            "warning": "#d29922",       # Orange
            "error": "#cb2431",         # Red
            "background": "#ffffff",     # White
            "surface": "#f5f5f5",       # Light gray
            "text": "#24292e",          # Almost black
            "text_secondary": "#57606a", # Dark gray
            "border": "#e0e0e0",        # Light gray
            "separator": "#d0d0d0"      # Slightly darker than border
        }
    }

    @staticmethod
    def get_colors():
        """Return color scheme based on appearance mode"""
        mode = ctk.get_appearance_mode()
        return AppTheme.COLORS["Dark" if mode == "Dark" else "Light"]

    @staticmethod
    def setup_theme():
        """Configure default theme settings"""
        def theme_change_event(*args):
            # Call all registered callbacks when theme changes
            mode = ctk.get_appearance_mode()
            for callback in AppTheme._theme_callbacks:
                try:
                    callback()
                except Exception as e:
                    print(f"Error in theme callback: {e}")

        if AppTheme._root is None:
            # Create root window if it doesn't exist
            AppTheme._root = ctk.CTk()
            AppTheme._root.withdraw()  # Hide the window
            AppTheme._root.protocol("WM_DELETE_WINDOW", AppTheme.cleanup)  # Handle window close
            
        # Set initial theme
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        
        # Bind theme change event
        AppTheme._root.bind("<<AppearanceModeChanged>>", theme_change_event)

        
    @staticmethod
    def add_theme_callback(callback):
        """Add a callback function that will be called when theme changes"""
        if callback not in AppTheme._theme_callbacks:
            AppTheme._theme_callbacks.append(callback)
    
    @staticmethod
    def remove_theme_callback(callback):
        """Remove a theme change callback"""
        if callback in AppTheme._theme_callbacks:
            AppTheme._theme_callbacks.remove(callback)

    @staticmethod
    def get_button_style(font=None, override_colors=None):
        """Return common button styling"""
        colors = AppTheme.get_colors()
        style = {
            "font": font if font else AppTheme.FONTS["normal"],
            "corner_radius": 6,
            "border_width": 0,
            "hover": True
        }
        
        # Only add colors if not being overridden
        if not override_colors:
            style.update({
                "fg_color": colors["primary"],
                "hover_color": colors["secondary"],
                "text_color": "#ffffff"
            })
            
        return style

    @staticmethod
    def get_frame_style(override_colors=None):
        """Return common frame styling"""
        colors = AppTheme.get_colors()
        style = {
            "corner_radius": 10,
            "border_width": 1
        }
        
        if not override_colors:
            style.update({
                "border_color": colors["border"],
                "fg_color": colors["surface"]
            })
            
        return style

    @staticmethod
    def get_label_style(font=None, is_secondary=False):
        """Return common label styling"""
        colors = AppTheme.get_colors()
        return {
            "font": font if font else AppTheme.FONTS["normal"],
            "text_color": colors["text_secondary"] if is_secondary else colors["text"]
        }

    @staticmethod
    def get_input_style():
        """Return common input styling"""
        colors = AppTheme.get_colors()
        return {
            "font": AppTheme.FONTS["normal"],
            "corner_radius": 6,
            "border_width": 1,
            "border_color": colors["border"],
            "fg_color": colors["surface"],
            "text_color": colors["text"]
        }

    @staticmethod
    def get_progress_style():
        """Return progress bar styling"""
        colors = AppTheme.get_colors()
        return {
            "corner_radius": 6,
            "border_width": 0,
            "fg_color": colors["surface"],
            "progress_color": colors["primary"]
        }

    @staticmethod
    def get_textbox_style():
        """Return textbox styling"""
        colors = AppTheme.get_colors()
        return {
            "font": AppTheme.FONTS["monospace"],
            "corner_radius": 6,
            "border_width": 1,
            "border_color": colors["border"],
            "fg_color": colors["surface"],
            "text_color": colors["text"]
        }

    @staticmethod
    def get_separator_color():
        """Return separator color based on theme"""
        return AppTheme.get_colors()["separator"]
    
    @staticmethod
    def cleanup():
        """Clean up resources when app closes"""
        if AppTheme._root:
            AppTheme._root.unbind("<<AppearanceModeChanged>>")
            AppTheme._root.destroy()
            AppTheme._root = None
        AppTheme._theme_callbacks.clear()

    @staticmethod
    def add_theme_callback(callback):
        """Add a callback function that will be called when theme changes"""
        if callback and callback not in AppTheme._theme_callbacks:
            AppTheme._theme_callbacks.append(callback)
    
    @staticmethod
    def remove_theme_callback(callback):
        """Remove a theme change callback"""
        if callback in AppTheme._theme_callbacks:
            AppTheme._theme_callbacks.remove(callback)
    
   