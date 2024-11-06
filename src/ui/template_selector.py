"""GUI for IMS template selection.

This module provides a graphical user interface for selecting and processing
IMS templates.

TODO:
- Add progress indicators for long operations
- Add template preview functionality
- Add error handling and recovery
- Add support for bulk processing
- Add template validation before processing
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from ..config.settings import TEMPLATE_DIR, WINDOW_TITLE, WINDOW_SIZE
from ..core.excel_processor import ExcelProcessor, save_to_json

class IMSSelector:
    """GUI for selecting and processing IMS templates."""
    
    def __init__(self):
        """Initialize the GUI window."""
        self.root = tk.Tk()
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self._setup_ui()
        
    def _setup_ui(self) -> None:
        """Setup UI components."""
        self.center_window()
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        ttk.Label(main_frame, 
                 text="Select IMS Template:", 
                 font=('Helvetica', 12)).grid(row=0, column=0, pady=10)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, pady=10)
        
        self.template_options = self.get_template_options()
        self._create_template_buttons(button_frame)

    def _create_template_buttons(self, frame: ttk.Frame) -> None:
        """Create buttons for each template option."""        
        for i, template in enumerate(self.template_options):
            ttk.Button(frame, 
                      text=template,
                      command=lambda t=template: self.select_template(t)).grid(
                          row=i, column=0, pady=5, padx=10, sticky='ew')
                    
        # Center the window
        self.center_window()
        
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Add label
        ttk.Label(main_frame, text="Select IMS Template:", font=('Helvetica', 12)).grid(row=0, column=0, pady=10)
        
        # Create frame for buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, pady=10)
        
        # Get template options
        self.template_options = self.get_template_options()
        
        # Create buttons for each template
        for i, template in enumerate(self.template_options):
            ttk.Button(button_frame, 
                      text=template,
                      command=lambda t=template: self.select_template(t)).grid(
                          row=i, column=0, pady=5, padx=10, sticky='ew')
    
    def center_window(self):
        """Center the window on the screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def get_template_options(self):
        """Get list of template folders"""
        # Use TEMPLATE_DIR from settings 
        template_dir = TEMPLATE_DIR / "IMS_TEMPLATE_COPIES"
        try:
            # Get list of subdirectories if directory exists
            if template_dir.exists():
                return [d.name for d in template_dir.iterdir() 
                        if d.is_dir()]
            return []
                    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read template directory: {e}")
            return []
    
    def select_template(self, template):
        """Handle template selection"""
        try:
            # Construct path to template directory
            template_dir = TEMPLATE_DIR / "IMS_TEMPLATE_COPIES" / template
            
            # Open file dialog in the specific template directory
            filename = filedialog.askopenfilename(
                title=f"Select {template} Template",
                initialdir=template_dir,
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )
            
            if not filename:  # User cancelled
                return
                
            template_path = Path(filename)
            
            # Check if file is .xls and convert if needed
            if template_path.suffix.lower() == '.xls':
                from ..core.excel_converter import ExcelConverter
                converter = ExcelConverter()
                xlsx_path = template_path.with_suffix('.xlsx')
                converter.convert_file(template_path, xlsx_path)
                template_path = xlsx_path
                
            # Process the template
            processor = ExcelProcessor(template_path)
            data = processor.process_workbook()
            
            # Save JSON in same directory as input file
            output_path = template_path.parent / f"{template_path.stem}.json"
            save_to_json(data, output_path)
            
            # Close template selector window
            self.root.destroy()
            
            # Launch task completion UI
            from .task_completion import TaskCompletionUI
            task_ui = TaskCompletionUI(output_path)
            task_ui.run()
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process template: {e}")
    
    def run(self):
        """Start the UI"""
        if not self.template_options:
            messagebox.showerror("Error", "No templates found in IMS_TEMPLATE_COPIES directory")
            self.root.destroy()
            return
        self.root.mainloop()

if __name__ == "__main__":
    app = IMSSelector()
    app.run()
