#!/usr/bin/env python3
"""
Build script for creating Pic2Doc executable
Builds a standalone executable for Mac using PyInstaller
"""

import PyInstaller.__main__
import platform
import sys
from pathlib import Path


def build_executable():
    """Build platform-specific executable"""

    system = platform.system()

    if system != 'Darwin':
        print(f"⚠ Warning: This build script is optimized for macOS.")
        print(f"  Current system: {system}")
        print(f"  Build may not work as expected.")
        print()

    print("=" * 70)
    print("Building Pic2Doc executable")
    print("=" * 70)
    print()

    # Common options
    options = [
        'src/gui_main.py',  # GUI version
        '--name=Pic2Doc',
        '--onefile',
        '--windowed',  # No console window for GUI
        '--clean',
        '--noconfirm',
        '--paths=src',  # Add src directory to Python path
    ]

    # macOS-specific options
    if system == 'Darwin':
        print("Building for macOS...")
        icon_path = Path('assets/Pic2Doc.icns')
        if icon_path.exists():
            options.append(f'--icon={icon_path}')
            print(f"  Using icon: {icon_path}")
        else:
            print("  ⚠ Icon not found, building without icon")

    # Add VERSION file as data
    version_file = Path('VERSION')
    if version_file.exists():
        options.append(f'--add-data={version_file}:.')
        print(f"  Including VERSION file")

    # Hidden imports for dependencies
    options.extend([
        '--hidden-import=openpyxl',
        '--hidden-import=openpyxl.cell._writer',
        '--hidden-import=docx',
        '--hidden-import=docx.oxml',
        '--hidden-import=customtkinter',
        '--hidden-import=PIL',
        '--hidden-import=PIL._tkinter_finder',
    ])

    print(f"Build options:")
    for opt in options:
        print(f"  {opt}")
    print()

    try:
        PyInstaller.__main__.run(options)
        print()
        print("=" * 70)
        print("✓ Build complete!")
        print("=" * 70)
        print()
        print(f"Executable location: dist/Pic2Doc")
        print()
        print("To test the executable:")
        print("  ./dist/Pic2Doc")
        print()
    except Exception as e:
        print()
        print("=" * 70)
        print("✗ Build failed!")
        print("=" * 70)
        print(f"\nError: {e}")
        sys.exit(1)


if __name__ == '__main__':
    build_executable()
