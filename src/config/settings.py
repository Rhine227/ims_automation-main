"""Configuration settings for the IMS Automation application."""
from pathlib import Path

# Directory Settings
BASE_DIR = Path(__file__).parent.parent.parent  # Changed to point to ims_automation folder
TEMPLATE_DIR = BASE_DIR / "data" / "templates"  # This will now point to ims_automation/data/templates
OUTPUT_DIR = BASE_DIR / "output"

# Excel Processing Settings
INPUT_HEADERS = {
    "Date Inspected by who?", 
    "OK", 
    "OK?", 
    "Date Inspected by who",
    "What to look for:",
    "What to look for?"
}
CATEGORY_COLOR = 'FFFFFF00'  # Yellow background
DEFAULT_VALUE = "no input"

# UI Settings
WINDOW_TITLE = "IMS Template Selector"
WINDOW_SIZE = "400x300"