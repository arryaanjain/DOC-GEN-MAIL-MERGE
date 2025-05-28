import customtkinter as ctk
from src.components.processing_panel import ProcessingPanel
from src.utils.config import AppConfig
from src.utils.theme import AppTheme

# Import DocumentProcessor only where needed
class MainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configure window
        self.title(f"{AppConfig.APP_NAME} v{AppConfig.APP_VERSION}")
        self.geometry(f"{AppConfig.WINDOW_WIDTH}x{AppConfig.WINDOW_HEIGHT}")
        self.minsize(AppConfig.WINDOW_MIN_WIDTH, AppConfig.WINDOW_MIN_HEIGHT)
        
        # Setup theme
        AppTheme.setup_theme()
        
        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Initialize document processor
        from src.backend.app.core.document_processor_new import DocumentProcessor
        self.doc_processor = DocumentProcessor(debug=AppConfig.DEBUG_MODE)
        
        # Create a scrollable frame for the processing panel
        self.scrollable_frame = ctk.CTkScrollableFrame(self)
        self.scrollable_frame.grid(
            row=0, column=0,
            padx=20, pady=20,
            sticky="nsew"
        )

        # Enable mouse wheel scrolling
        self.scrollable_frame.bind_all("<MouseWheel>", lambda event: self.scrollable_frame._parent_canvas.yview_scroll(int(-1*(event.delta/120)), "units"))
        self.scrollable_frame.bind_all("<Button-4>", lambda event: self.scrollable_frame._parent_canvas.yview_scroll(-1, "units"))  # For Linux
        self.scrollable_frame.bind_all("<Button-5>", lambda event: self.scrollable_frame._parent_canvas.yview_scroll(1, "units"))   # For Linux

        # Place ProcessingPanel inside the scrollable frame
        self.processing_panel = ProcessingPanel(
            self.scrollable_frame,
            process_callback=self.process_document
        )
        self.processing_panel.pack(fill="both", expand=True)

    def process_document(self, input_path: str, output_path: str, processing_date: str, append_file_path=None):
        """Callback for document processing"""
        return self.doc_processor.convert_docx_to_xlsx(
            docx_file_path=input_path,
            xlsx_file_path=output_path,
            processing_date=processing_date,
            append_to_file=append_file_path
        )