"""Main entry point for IMS Automation application.

This module serves as the entry point for the IMS (Inspection and Maintenance System) 
automation tool. It initializes the application and launches the template selector GUI.

TODO: Add error handling and logging configuration
"""

import os
import sys

# Add project root to Python path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

try:
    from src.ui.template_selector import IMSSelector
except ImportError as e:
    print(f"Error importing required modules: {e}")
    print(f"Python path: {sys.path}")
    sys.exit(1)

def main():
    """Launch the IMS Template Selector application."""
    try:
        app = IMSSelector()
        app.run()
    except Exception as e:
        print(f"Error running application: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()