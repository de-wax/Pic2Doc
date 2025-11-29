"""
Application constants for Pic2Doc
"""

# Configuration file name
CONFIG_FILE = "pic2doc_config.json"

# Default configuration values
DEFAULT_CONFIG = {
    'excel_file': 'beschreibungen.xlsx',
    'image_folder': 'pics',
    'output_file': 'output_document.docx',
    'filename_column': 'A',
    'caption_columns': ['I'],  # List of columns for captions
    'caption_separator': ' - ',  # Separator for multi-column captions
    'images_per_page': 3,
    'font_name': 'Arial',
    'font_size': 10,
    'font_bold': False,
    'font_italic': False,
    'font_underline': False,
    # v0.2.0 features
    'test_mode': False,
    'test_image_limit': 10,
    'margin_top_cm': 1.27,     # Word default: 1.27cm (0.5 inch)
    'margin_bottom_cm': 1.27,
    'margin_left_cm': 1.27,
    'margin_right_cm': 1.27,
    'smart_layout': True,      # Always enabled: intelligent side-by-side layout
}

# Supported image extensions
SUPPORTED_IMAGE_EXTENSIONS = ['.jpg', '.jpeg', '.png', '.bmp']

# Image orientation thresholds
LANDSCAPE_RATIO = 1.2  # width/height > 1.2 = landscape
PORTRAIT_RATIO = 0.8   # width/height < 0.8 = portrait
# Between 0.8 and 1.2 = square/neutral
