#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pic2Doc - Image to Document Generator
Creates Word documents from images with captions from Excel data

Author: René
Version: 0.1.0
"""

import sys
import os
from pathlib import Path

# Add parent directory to path to allow imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.core.config_manager import ConfigManager
from src.core.excel_reader import ExcelReader
from src.core.image_handler import ImageHandler
from src.core.document_generator import DocumentGenerator
from src.utils.constants import DEFAULT_CONFIG


def input_with_default(prompt: str, default):
    """
    Prompt for input with default value

    Args:
        prompt: Question to display
        default: Default value if nothing entered

    Returns:
        User input or default value
    """
    user_input = input(f"{prompt} [{default}]: ").strip()
    return user_input if user_input else default


def input_yes_no(prompt: str, default: bool = False) -> bool:
    """
    Prompt for yes/no input

    Args:
        prompt: Question to display
        default: Default value (True=Yes, False=No)

    Returns:
        True for yes, False for no
    """
    default_str = "J/n" if default else "j/N"
    user_input = input(f"{prompt} [{default_str}]: ").strip().lower()

    if not user_input:
        return default

    return user_input in ['j', 'ja', 'y', 'yes']


def get_user_configuration(saved_config: dict) -> dict:
    """
    Prompt user for all configuration parameters

    Args:
        saved_config: Saved or default configuration

    Returns:
        Dictionary with configuration values
    """
    print("=" * 70)
    print("KONFIGURATION")
    print("=" * 70)
    print("Drücke Enter um Standardwerte zu übernehmen")
    print("(Deine letzten Einstellungen werden als Standard verwendet)")
    print()

    config = {}

    # Files and folders
    print("--- Dateien und Ordner ---")
    config['excel_file'] = input_with_default(
        "Excel-Datei mit Beschreibungen",
        saved_config['excel_file']
    )

    config['image_folder'] = input_with_default(
        "Ordner mit Bildern",
        saved_config['image_folder']
    )

    config['output_file'] = input_with_default(
        "Name der Ausgabedatei",
        saved_config['output_file']
    )

    print()

    # Layout settings
    print("--- Layout ---")
    while True:
        try:
            images_per_page = input_with_default(
                "Anzahl Bilder pro Seite",
                str(saved_config['images_per_page'])
            )
            config['images_per_page'] = int(images_per_page)
            if config['images_per_page'] < 1:
                print("⚠ Bitte mindestens 1 Bild pro Seite angeben!")
                continue
            break
        except ValueError:
            print("⚠ Bitte eine gültige Zahl eingeben!")

    print()

    # Font settings
    print("--- Schriftformatierung für Bildunterschriften ---")

    config['font_name'] = input_with_default(
        "Schriftart",
        saved_config['font_name']
    )

    while True:
        try:
            font_size = input_with_default(
                "Schriftgröße (Punkt)",
                str(saved_config['font_size'])
            )
            config['font_size'] = int(font_size)
            if config['font_size'] < 1:
                print("⚠ Schriftgröße muss größer als 0 sein!")
                continue
            break
        except ValueError:
            print("⚠ Bitte eine gültige Zahl eingeben!")

    config['font_bold'] = input_yes_no(
        "Fettdruck",
        saved_config['font_bold']
    )

    config['font_italic'] = input_yes_no(
        "Kursiv",
        saved_config['font_italic']
    )

    config['font_underline'] = input_yes_no(
        "Unterstrichen",
        saved_config['font_underline']
    )

    print()

    # Advanced settings (v0.2.0)
    print("--- Erweiterte Einstellungen ---")

    # Test mode
    config['test_mode'] = input_yes_no(
        "Test-Modus (nur wenige Bilder verarbeiten)",
        saved_config.get('test_mode', False)
    )

    if config['test_mode']:
        while True:
            try:
                test_limit = input_with_default(
                    "Anzahl Bilder im Test-Modus",
                    str(saved_config.get('test_image_limit', 10))
                )
                config['test_image_limit'] = int(test_limit)
                if config['test_image_limit'] < 1:
                    print("⚠ Bitte mindestens 1 Bild angeben!")
                    continue
                break
            except ValueError:
                print("⚠ Bitte eine gültige Zahl eingeben!")
    else:
        config['test_image_limit'] = saved_config.get('test_image_limit', 10)

    # Smart layout is always enabled (v0.3.0)
    config['smart_layout'] = True

    # Caption columns configuration
    print()
    print("--- Bildunterschrift-Spalten ---")

    # Get saved caption columns
    saved_caption_cols = saved_config.get('caption_columns', ['I'])
    if isinstance(saved_caption_cols, str):  # Backward compatibility
        saved_caption_cols = [saved_caption_cols]

    default_cols_str = ','.join(saved_caption_cols)
    caption_cols_input = input_with_default(
        "Bildunterschrift-Spalten (kommagetrennt, z.B. 'I' oder 'A,B,I')",
        default_cols_str
    )

    # Parse column input
    caption_columns = [col.strip().upper() for col in caption_cols_input.split(',') if col.strip()]
    config['caption_columns'] = caption_columns

    if len(caption_columns) > 1:
        separator_input = input_with_default(
            "Trenner für mehrere Spalten",
            saved_config.get('caption_separator', ' - ')
        )
        config['caption_separator'] = separator_input
    else:
        config['caption_separator'] = saved_config.get('caption_separator', ' - ')

    # Margins
    margin_config = input_yes_no(
        "Seitenränder anpassen",
        False
    )

    if margin_config:
        print("\n  Seitenränder (Standard: 1.27 cm):")
        for margin_name in ['top', 'bottom', 'left', 'right']:
            key = f'margin_{margin_name}_cm'
            while True:
                try:
                    margin = input_with_default(
                        f"    {margin_name.capitalize()}-Rand (cm)",
                        str(saved_config.get(key, 1.27))
                    )
                    config[key] = float(margin)
                    if config[key] < 0:
                        print("    ⚠ Rand darf nicht negativ sein!")
                        continue
                    break
                except ValueError:
                    print("    ⚠ Bitte eine gültige Zahl eingeben!")
    else:
        config['margin_top_cm'] = saved_config.get('margin_top_cm', 1.27)
        config['margin_bottom_cm'] = saved_config.get('margin_bottom_cm', 1.27)
        config['margin_left_cm'] = saved_config.get('margin_left_cm', 1.27)
        config['margin_right_cm'] = saved_config.get('margin_right_cm', 1.27)

    # Filename column (fixed for now)
    config['filename_column'] = 'A'

    print()

    return config


def display_configuration(config: dict):
    """
    Display current configuration

    Args:
        config: Configuration dictionary
    """
    print("=" * 70)
    print("GEWÄHLTE KONFIGURATION")
    print("=" * 70)
    print(f"Excel-Datei:          {config['excel_file']}")
    print(f"Bilder-Ordner:        {config['image_folder']}")
    print(f"Ausgabedatei:         {config['output_file']}")
    print()
    print(f"Dateinamen-Spalte:    {config['filename_column']}")

    caption_cols = config.get('caption_columns', ['I'])
    caption_cols_str = ', '.join(caption_cols)
    print(f"Bildunterschrift-Spalten: {caption_cols_str}")
    if len(caption_cols) > 1:
        print(f"  Trenner: '{config.get('caption_separator', ' - ')}'")

    print()
    print(f"Bilder pro Seite:     {config['images_per_page']}")
    print(f"  (Intelligente Anordnung: nebeneinander wenn möglich)")
    print()
    print(f"Schriftart:           {config['font_name']}")
    print(f"Schriftgröße:         {config['font_size']} pt")
    print(f"Fett:                 {'Ja' if config['font_bold'] else 'Nein'}")
    print(f"Kursiv:               {'Ja' if config['font_italic'] else 'Nein'}")
    print(f"Unterstrichen:        {'Ja' if config['font_underline'] else 'Nein'}")
    print("=" * 70)
    print()


def main():
    """Main entry point"""
    print()
    print("=" * 70)
    print(" " * 20 + "PIC2DOC")
    print("=" * 70)
    print()
    print("Erstellt formatierte Word-Dokumente aus Bildern und Excel-Beschreibungen")
    print()

    # Load saved configuration
    config_manager = ConfigManager()
    saved_config = config_manager.load_config()
    print()

    # Get configuration from user
    config = get_user_configuration(saved_config)

    # Display configuration
    display_configuration(config)

    # Confirm
    if not input_yes_no("Mit dieser Konfiguration fortfahren?", True):
        print("\nAbgebrochen.")
        return

    print()
    print("=" * 70)
    print("VERARBEITUNG")
    print("=" * 70)
    print()

    # Validate files exist
    if not os.path.exists(config['excel_file']):
        print(f"✗ Fehler: Excel-Datei nicht gefunden: {config['excel_file']}")
        print(f"\nErwartetes Format:")
        print(f"  Spalte A: Dateiname ohne Endung (z.B. 'FIAS 21B 000001')")
        print(f"  Spalte I: Beschreibung")
        return

    if not os.path.exists(config['image_folder']):
        print(f"✗ Fehler: Bilder-Ordner nicht gefunden: {config['image_folder']}")
        return

    # Read Excel data
    try:
        excel_reader = ExcelReader()
        excel_data = excel_reader.read_data(
            config['excel_file'],
            config['filename_column'],
            config.get('caption_columns', ['I']),
            config.get('caption_separator', ' - ')
        )
    except Exception as e:
        print(f"✗ Fehler beim Lesen der Excel-Datei: {e}")
        return

    if not excel_data:
        print("✗ Keine Daten in Excel-Datei gefunden!")
        print("\nStelle sicher, dass:")
        print("  - Die Datei nicht leer ist")
        print(f"  - Dateinamen in Spalte {config['filename_column']} stehen")
        caption_cols_str = ', '.join(config.get('caption_columns', ['I']))
        print(f"  - Beschreibungen in Spalte(n) {caption_cols_str} stehen")
        return

    # Apply test mode limit if enabled
    if config.get('test_mode', False):
        test_limit = config.get('test_image_limit', 10)
        excel_data = excel_data[:test_limit]
        print(f"⚡ Test-Modus aktiv: Verarbeite nur die ersten {len(excel_data)} Bilder")
        print()

    # Validate and locate images
    try:
        image_handler = ImageHandler(config['image_folder'])

        # Build complete image data with paths and orientation info
        complete_data = []
        smart_layout = config.get('smart_layout', False)

        for filename, caption in excel_data:
            try:
                if smart_layout:
                    # Get detailed image info including orientation
                    image_info = image_handler.get_image_info(filename)
                    complete_data.append((filename, caption, image_info['path'], image_info))
                else:
                    # Just get path (old behavior)
                    image_path = image_handler.get_image_path(filename)
                    complete_data.append((filename, caption, image_path, None))
            except (FileNotFoundError, ValueError) as e:
                print(f"⚠ {e}")
                continue

        if not complete_data:
            print("✗ Keine Bilder gefunden!")
            return

        print(f"✓ {len(complete_data)} Bilder gefunden und validiert")

        # Show orientation stats if smart layout is enabled
        if smart_layout:
            orientations = {}
            for _, _, _, info in complete_data:
                if info:
                    ori = info['orientation']
                    orientations[ori] = orientations.get(ori, 0) + 1

            if orientations:
                print(f"  Orientierungen: ", end="")
                print(", ".join([f"{count} {ori}" for ori, count in orientations.items()]))

        print()

    except Exception as e:
        print(f"✗ Fehler bei der Bildverarbeitung: {e}")
        return

    # Generate document
    try:
        doc_generator = DocumentGenerator(config)
        processed, errors = doc_generator.create_document(
            complete_data,
            config['output_file']
        )
    except Exception as e:
        print(f"\n✗ Fehler beim Erstellen des Dokuments: {e}")
        import traceback
        traceback.print_exc()
        return

    # Save configuration
    print()
    config_manager.save_config(config)

    print("\n✓ Fertig!")
    print("\nDeine Einstellungen wurden gespeichert und werden beim")
    print("nächsten Start als Standardwerte verwendet.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nAbgebrochen durch Benutzer.")
        sys.exit(0)
    except Exception as e:
        print(f"\n✗ Unerwarteter Fehler: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
