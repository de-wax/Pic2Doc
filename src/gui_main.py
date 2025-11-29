#!/usr/bin/env python3
"""
Pic2Doc GUI Entry Point
Launches the graphical user interface
"""

import sys
from pathlib import Path

# Add src directory to path for imports
src_dir = Path(__file__).parent
sys.path.insert(0, str(src_dir))

if __name__ == "__main__":
    from gui.main_window import main
    main()
