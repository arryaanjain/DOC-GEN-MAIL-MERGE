import sys
import os
from pathlib import Path

def setup_paths():
    """Setup Python path to include all necessary directories"""
    # Get absolute paths
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent
    backend_dir = project_root / 'backend'
    
    # Print paths for debugging
    print("Setting up Python paths:")
    print(f"Current dir: {current_dir}")
    print(f"Project root: {project_root}")
    print(f"Backend dir: {backend_dir}")
    
    # Add backend directory first (higher priority)
    if backend_dir.exists():
        sys.path.insert(0, str(backend_dir))
        print("Added backend directory to path")
    else:
        raise FileNotFoundError(f"Backend directory not found at {backend_dir}")
        
    # Add desktop directory for src imports
    sys.path.insert(0, str(current_dir))
    print("Added desktop directory to path")
    
    # Verify backend package structure
    backend_init = backend_dir / 'app' / '__init__.py'
    core_init = backend_dir / 'app' / 'core' / '__init__.py'
    
    # Create __init__.py files if they don't exist
    backend_init.parent.mkdir(parents=True, exist_ok=True)
    core_init.parent.mkdir(parents=True, exist_ok=True)
    
    if not backend_init.exists():
        backend_init.touch()
        print(f"Created {backend_init}")
    if not core_init.exists():
        core_init.touch()
        print(f"Created {core_init}")

    # Print current Python path for verification
    print("\nPython path after setup:")
    for p in sys.path:
        print(f"- {p}")


from src.views.main_window import MainWindow
from src.utils.config import AppConfig

def setup_environment():
    """Setup necessary directories and environment variables"""
    # Create default directories if they don't exist
    os.makedirs(AppConfig.DEFAULT_OUTPUT_DIR, exist_ok=True)
    os.makedirs(AppConfig.DEFAULT_LOG_DIR, exist_ok=True)

if __name__ == "__main__":
    setup_paths()
    
    # Now import required modules after path setup
    import os
    from src.views.main_window import MainWindow
    from src.utils.config import AppConfig
    
    try:
        os.makedirs(AppConfig.DEFAULT_OUTPUT_DIR, exist_ok=True)
        os.makedirs(AppConfig.DEFAULT_LOG_DIR, exist_ok=True)
        
        app = MainWindow()
        app.mainloop()
    except Exception as e:
        print(f"Application error: {str(e)}")
        sys.exit(1)