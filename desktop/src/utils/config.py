class AppConfig:
    # Application settings
    APP_NAME = "Document Processor"
    APP_VERSION = "1.0.0"
    
    # Window settings
    WINDOW_WIDTH = 800
    WINDOW_HEIGHT = 600
    WINDOW_MIN_WIDTH = 600
    WINDOW_MIN_HEIGHT = 400
    
    # File settings
    SUPPORTED_FORMATS = [
        ("Word Documents", "*.docx"),
        ("All Files", "*.*")
    ]
    
    # Processing settings
    DEBUG_MODE = True
    MAX_LOG_ENTRIES = 1000
    
    # Default paths
    DEFAULT_OUTPUT_DIR = "outputs"
    DEFAULT_LOG_DIR = "logs"

    # Theme settings
    DEFAULT_THEME_MODE = "system"  # system, light, dark
    DEFAULT_COLOR_THEME = "blue"

    # Grid settings
    GRID_PADDING = {
        "padx": 10,
        "pady": 5
    }