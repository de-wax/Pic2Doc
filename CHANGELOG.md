# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.4.2] - 2025-11-29

### Fixed
- Settings now properly persist across all application restarts
- Configuration file is created immediately on first launch
- Settings are saved automatically when selecting files/folders
- Theme changes are saved immediately
- Duplicate values no longer inserted when loading saved config

### Changed
- Auto-save enabled for all user interactions (file selection, theme changes)
- Improved settings loading to prevent duplicate entries

## [0.4.1] - 2025-11-29

### Added
- Settings persistence on window close
- Theme preference saved and restored between sessions
- macOS foreground activation on app launch

### Fixed
- Settings not being loaded on first application start
- Document creation bug when test mode is disabled
- Application not appearing in foreground when launched via double-click on macOS
- Crash when reading test limit value with test mode disabled

### Changed
- Configuration now automatically saves when closing the application
- All settings from last session are restored on startup

## [0.4.0] - 2025-11-28

### Added
- **GUI Application**: Modern graphical interface using CustomTkinter
- **Theme Support**: System/Light/Dark theme selector in title bar
- **File Dialogs**: Native file and folder selection dialogs
- **Progress Tracking**: Real-time progress bar with current file display
- **Cancel Functionality**: Ability to cancel document generation in progress
- **Error Display Panel**: Scrollable error log showing all processing errors
- **File Overwrite Warning**: Dialog confirmation before overwriting existing files
- **Multi-column Captions**: Support for combining multiple Excel columns with configurable separator
- **Enhanced Error Reporting**: Track filename AND error reason for up to 10 failures
- **GUI Entry Point**: `src/gui_main.py` for launching GUI application
- **macOS App Bundle**: `dist/Pic2Doc.app` for double-click launching
- **EXAMPLE_DATA.md**: Comprehensive data format documentation

### Changed
- Build script now creates windowed GUI application instead of CLI
- Default entry point changed from `main.py` to `gui_main.py`
- Excel filename column now expects WITH extension (was WITHOUT in v0.1.0)
- Caption columns configuration changed from single column to list
- Configuration schema extended with `caption_separator` and list-based `caption_columns`
- Window size optimized to 750x750px for better space utilization
- Browse buttons shortened to "..." (40px) for cleaner layout
- Combined Layout & Font settings into single section
- Removed redundant smart layout info text from GUI

### Fixed
- **Caption Wrapping**: Implemented cantSplit properties to prevent table breaks across pages
- **Grid Layout**: Intelligent grid calculation for optimal image arrangement (e.g., 3 images → 2×2 grid)
- **Conservative Sizing**: Images 10% smaller with increased overhead calculations
- **Page Breaks**: Single table per page prevents unwanted breaks
- Module import issues in PyInstaller build with `--paths=src` option

### Removed
- Manual image dimension settings (now fully automatic based on grid)
- Fixed width/height configuration options
- Margins configuration dialog from GUI (uses standard 1.27cm margins)
- Smart layout toggle (always enabled)

## [0.3.0] - 2025-11-26

### Added
- **Intelligent Grid-Based Layout**: Automatic optimal grid calculation for images
- **Automatic Image Sizing**: Images sized to fit grid based on images_per_page
- **Table Keep-Together**: XML properties prevent caption wrapping to next page
- **Multi-Column Caption Support**: Select multiple Excel columns with configurable separator
- **Enhanced Error Reporting**: Display up to 10 errors with specific messages
- **Conservative Page Height**: Better space calculations to guarantee fit

### Changed
- Layout system now uses single table per page with all images
- Caption spacing calculated as `(font_size / 72) * 2.0` inches per row
- Images arranged both horizontally AND vertically in grid
- Grid examples: 3→2×2, 4→2×2, 5→3×2, 6→3×2, 8→3×3
- Default configuration updated with `caption_columns` list and `caption_separator`

### Fixed
- Caption wrapping issue resolved with conservative calculations
- Empty pages eliminated with proper grid layout
- Images now respect strict Excel sheet order

## [0.2.0] - 2025-11-25

### Added
- **Test mode**: Process limited number of images for quick testing (configurable limit)
- **Smart layout**: Automatic image orientation detection (portrait/landscape/square)
- **Intelligent sizing**: Portrait images rendered narrower (75%), landscape at full width (100%), square at 85%
- **Configurable margins**: Adjust document margins (top, bottom, left, right) in cm
- **Image orientation stats**: Display breakdown of portrait/landscape/square images when smart layout is enabled
- **Pillow dependency**: Added for image dimension analysis and orientation detection
- Project-local `.claude/CLAUDE.md` for development tracking and notes

### Changed
- DocumentGenerator now accepts optional image info dictionary for smart layout
- Configuration schema extended with: `test_mode`, `test_image_limit`, `margin_*_cm`, `smart_layout`
- ImageHandler enhanced with orientation detection methods
- CLI now includes "Erweiterte Einstellungen" section for new features
- Updated requirements.txt to explicitly list Pillow>=10.0.0

### Improved
- Better space utilization with configurable margins (default remains 1.27cm)
- Mixed orientation images now render more appropriately on same page
- Faster iteration during development with test mode

## [0.1.0] - 2025-11-25

### Added
- Initial refactored version of Pic2Doc
- Properly structured Python project with separated core modules
- Excel reading from column A (filenames without extension) and column I (captions)
- Image processing from pics/ folder (automatically adds .jpg extension)
- Word document generation with configurable layout and formatting
- Configuration management with JSON persistence
- English code with German CLI messages
- VERSION file for semantic versioning
- CHANGELOG.md for tracking changes
- Project standards compliance (SemVer, English code/comments)

### Changed
- Renamed project from "word_generator" to "Pic2Doc"
- Refactored monolithic script into modular architecture
- Translated all code, variable names, and comments from German to English
- Fixed Excel caption reading: Changed from column B to column I

### Fixed
- Critical bug: Excel reader now correctly reads captions from column I instead of column B
- Improved error handling and user feedback
- Better file path handling for cross-platform compatibility
