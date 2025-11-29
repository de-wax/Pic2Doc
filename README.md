# Pic2Doc - Image to Document Generator

Transform collections of images into professionally formatted Word documents with customizable captions from Excel data.

**Version:** 0.4.0
**Author:** René

## Features

### Core Functionality
- ✅ Batch process thousands of images efficiently
- ✅ Read image filenames and captions from Excel files
- ✅ Intelligent grid-based layout (automatic optimal arrangement)
- ✅ Multi-column caption support with configurable separator
- ✅ Automatic image sizing based on grid layout
- ✅ No caption wrapping - everything fits on one page

### User Interface
- ✅ Modern GUI with CustomTkinter
- ✅ Theme selector (System/Light/Dark)
- ✅ File selection dialogs
- ✅ Real-time progress tracking
- ✅ Cancel functionality
- ✅ Error display panel
- ✅ File overwrite warnings

### Technical
- ✅ Customizable font formatting (family, size, bold, italic, underline)
- ✅ Configurable margins
- ✅ Test mode for quick iteration
- ✅ Persistent configuration
- ✅ Standalone macOS application (no Python required)
- ✅ German user interface with English codebase

## Quick Start

### Using the Application (Recommended)

1. **Double-click** `Pic2Doc.app` (macOS)
2. **Select your files** using the GUI:
   - Excel file with image descriptions
   - Image folder
   - Output location for Word document
3. **Configure settings** (optional)
4. **Click "Dokument erstellen"**

### Running from Source

```bash
# Prerequisites: Python 3.8 or higher

# 1. Set up virtual environment
python3 -m venv venv
source venv/bin/activate  # On Mac/Linux

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the GUI application
python src/gui_main.py

# Or run the CLI version
python src/main.py
```

## Data Format

### Excel File Structure

See [EXAMPLE_DATA.md](EXAMPLE_DATA.md) for detailed examples.

Your Excel file should contain:
- **Column A** (or your chosen column): Image filenames WITH extension
- **Column I** (or your chosen column(s)): Caption text

Example:

| A              | B        | I      |
|----------------|----------|--------|
| photo001.jpg   | Nature   | N-001  |
| photo002.jpg   | Urban    | U-002  |
| photo003.jpg   | Wildlife | W-003  |

### Multi-Column Captions

Combine multiple columns for richer captions:
- Single column: `I` → "N-001"
- Multiple columns: `B,I` → "Nature - N-001"
- Separator is configurable (default: " - ")

### Image Folder

Create a folder containing all referenced images:
```
images/
├── photo001.jpg
├── photo002.jpg
└── photo003.jpg
```

Supported formats: JPEG, PNG, and other Pillow-compatible formats.

## Project Structure

```
Pic2Doc/
├── src/
│   ├── gui_main.py                # GUI entry point
│   ├── main.py                    # CLI entry point
│   ├── gui/
│   │   └── main_window.py         # GUI application
│   ├── core/
│   │   ├── excel_reader.py        # Excel file processing
│   │   ├── document_generator.py  # Word document creation
│   │   ├── image_handler.py       # Image file operations
│   │   └── config_manager.py      # Configuration management
│   └── utils/
│       └── constants.py            # Application constants
├── assets/                         # Application icons
├── dist/                           # Built executables
│   └── Pic2Doc.app                # macOS application
├── build.py                        # Build script
├── requirements.txt                # Python dependencies
├── VERSION                         # Version number
├── CHANGELOG.md                    # Version history
├── EXAMPLE_DATA.md                 # Data format examples
└── README.md                       # This file
```

## Configuration

The application stores settings in `pic2doc_config.json`:

```json
{
  "excel_file": "data.xlsx",
  "image_folder": "images",
  "output_file": "output_document.docx",
  "filename_column": "A",
  "caption_columns": ["I"],
  "caption_separator": " - ",
  "images_per_page": 3,
  "font_name": "Arial",
  "font_size": 10,
  "test_mode": false,
  "test_image_limit": 10
}
```

Configuration is automatically saved and loaded between sessions.

## Building from Source

### macOS Application

```bash
# 1. Activate virtual environment
source venv/bin/activate

# 2. Install build dependencies
pip install -r requirements.txt

# 3. Run build script
python build.py

# 4. Application created at dist/Pic2Doc.app
open dist/Pic2Doc.app
```

### Build Output

- `dist/Pic2Doc.app` - macOS application bundle (double-click to run)
- `dist/Pic2Doc` - Standalone executable (command-line runnable)

## Usage

### GUI Mode (Default)

1. Launch the application
2. **Input Files**:
   - Select Excel file with descriptions
   - Select folder containing images
   - Choose output location
3. **Configuration**:
   - Filename column (e.g., "A")
   - Caption columns (e.g., "I" or "B,I")
   - Images per page (1-10)
   - Font settings
4. **Optional**:
   - Enable test mode (process only N images)
   - Change theme (System/Light/Dark)
5. **Generate**: Click "Dokument erstellen"

### CLI Mode

```bash
python src/main.py
# Follow interactive prompts
```

## Features in Detail

### Intelligent Grid Layout

The application automatically calculates the optimal grid for your images:
- 3 images → 2×2 grid (one empty cell)
- 4 images → 2×2 grid
- 5 images → 3×2 grid (one empty cell)
- 6 images → 3×2 grid
- 8 images → 3×3 grid (one empty cell)

Images are sized to fit perfectly within the grid with captions.

### Caption Management

- **No wrapping**: Captions never wrap to the next page
- **Conservative spacing**: Automatic calculation ensures everything fits
- **Multi-column**: Combine data from multiple Excel columns
- **Customizable separator**: Default " - " or your choice

### Error Handling

- File overwrite warnings
- Missing image detection
- Error panel with detailed messages
- Up to 10 errors displayed with specific reasons

## Troubleshooting

### Application won't start
- **macOS**: Right-click → Open (first time only, due to security)
- **Permission denied**: `chmod +x dist/Pic2Doc.app/Contents/MacOS/Pic2Doc`

### "Excel file not found"
- Ensure file path is correct
- File must be `.xlsx` format

### "Images not found"
- Verify filenames in Excel match actual files exactly (case-sensitive)
- Include file extensions in Excel (e.g., "photo.jpg", not "photo")
- Check images exist in selected folder

### Captions don't appear
- Verify caption column letter is correct
- Check cells contain data
- Row 1 is treated as header and skipped

### Images don't fit properly
- Try reducing images per page
- Use test mode to experiment with layout
- Check image resolution (very large images may cause issues)

## Development

### Requirements

- Python 3.8 or higher
- See `requirements.txt` for dependencies

### Running Tests

```bash
# Test GUI
python src/gui_main.py

# Test CLI
python src/main.py

# Test with limited images
# Enable test mode in GUI or config
```

### Code Standards

- **Language**: English code/comments, German UI
- **Style**: PEP 8
- **Documentation**: Docstrings for all public functions
- **Versioning**: Semantic Versioning (SemVer)

## Version History

See [CHANGELOG.md](CHANGELOG.md) for detailed version history.

### Recent Releases

- **v0.4.0** - GUI with CustomTkinter, theme support, enhanced error handling
- **v0.3.0** - Grid-based layout, multi-column captions, no caption wrapping
- **v0.2.0** - Test mode, smart layout, configurable margins
- **v0.1.0** - Initial CLI version

## Roadmap

### Future Versions

- **v0.5.0**: Windows builds (.exe), cross-platform testing
- **v0.6.0**: PDF export option, batch processing multiple Excel files
- **v0.7.0**: Image preprocessing (resize, crop, filters)

## License

Free for personal and commercial use.

See [LICENSE](LICENSE) for details.

## Contributing

This is a personal project, but suggestions and bug reports are welcome via GitHub issues.

## Contact

Created by René

For German user documentation, see [README.de.md](README.de.md)

## Acknowledgments

Built with:
- [python-docx](https://python-docx.readthedocs.io/) - Word document generation
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel file reading
- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) - Modern GUI framework
- [Pillow](https://pillow.readthedocs.io/) - Image processing
