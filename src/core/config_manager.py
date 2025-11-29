"""
Configuration Manager for Pic2Doc
Handles loading, saving, and migrating configuration files
"""

import json
import os
import sys
from pathlib import Path
from typing import Dict, Any
from ..utils.constants import CONFIG_FILE, DEFAULT_CONFIG


class ConfigManager:
    """Manages application configuration persistence"""

    def __init__(self, config_path: str = None):
        """
        Initialize configuration manager

        Args:
            config_path: Optional custom path to config file
        """
        if config_path:
            self.config_path = config_path
        else:
            # When running as PyInstaller bundle, save config in user's home directory
            if getattr(sys, 'frozen', False):
                # Running as bundle - use home directory
                home_dir = Path.home()
                self.config_path = home_dir / CONFIG_FILE
                print(f"DEBUG: Running as bundle, config path: {self.config_path}")
            else:
                # Running normally - use current directory
                self.config_path = CONFIG_FILE
                print(f"DEBUG: Running from source, config path: {self.config_path}")

    def load_config(self) -> Dict[str, Any]:
        """
        Load configuration from file

        Returns:
            Dictionary with configuration values
        """
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                print(f"✓ Konfiguration geladen aus: {self.config_path}")

                # Migrate old config format if needed
                config = self._migrate_config(config)

                return config
            except Exception as e:
                print(f"⚠ Fehler beim Laden der Konfiguration: {e}")
                print("  Verwende Standard-Konfiguration")
                return DEFAULT_CONFIG.copy()
        else:
            return DEFAULT_CONFIG.copy()

    def save_config(self, config: Dict[str, Any]) -> bool:
        """
        Save configuration to file

        Args:
            config: Dictionary with configuration values

        Returns:
            True if successful, False otherwise
        """
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            print(f"✓ Konfiguration gespeichert in: {self.config_path}")
            return True
        except Exception as e:
            print(f"⚠ Fehler beim Speichern der Konfiguration: {e}")
            return False

    def _migrate_config(self, config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Migrate old configuration format to new format

        Args:
            config: Old configuration dictionary

        Returns:
            Migrated configuration dictionary
        """
        # Check if this is an old format (from word_generator)
        migrated = False

        # Add caption_column if missing (new in v0.1.0)
        if 'caption_column' not in config:
            config['caption_column'] = 'I'
            migrated = True

        # Add filename_column if missing
        if 'filename_column' not in config:
            config['filename_column'] = 'A'
            migrated = True

        # Ensure all default keys exist
        for key, value in DEFAULT_CONFIG.items():
            if key not in config:
                config[key] = value
                migrated = True

        if migrated:
            print("  → Konfiguration wurde auf neues Format aktualisiert")

        return config
