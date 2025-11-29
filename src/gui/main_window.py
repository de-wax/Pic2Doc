"""
Pic2Doc GUI - Main Window
Modern GUI for image-to-document conversion
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
from pathlib import Path
import sys
import os

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from src.core.config_manager import ConfigManager
from src.core.excel_reader import ExcelReader
from src.core.image_handler import ImageHandler
from src.core.document_generator import DocumentGenerator


class Pic2DocGUI(ctk.CTk):
    """Main GUI window for Pic2Doc"""

    def __init__(self):
        super().__init__()

        # Window setup
        self.title("Pic2Doc")
        self.geometry("750x750")
        self.resizable(True, True)

        # Set theme
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        # Load configuration
        self.config_manager = ConfigManager()
        self.config = self.config_manager.load_config()

        # Processing state
        self.is_processing = False
        self.cancel_processing = False
        self.processing_thread = None
        self.error_list = []  # Store errors during processing

        # Loading state - prevent auto-save during initial load
        self.is_loading = True

        # Save settings on close
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Create GUI
        self.create_widgets()
        self.load_saved_config()

        # Enable auto-save after initial load
        self.is_loading = False

        # Bring to foreground on macOS
        self.bring_to_foreground()

    def create_widgets(self):
        """Create all GUI widgets"""

        # Main container with padding
        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=20, pady=20)

        # ===== TITLE BAR WITH THEME SELECTOR =====
        title_bar = ctk.CTkFrame(main_container, fg_color="transparent")
        title_bar.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(title_bar, text="Pic2Doc", font=("Arial", 24, "bold")).pack(side="left")

        # Theme selector on the right
        theme_frame = ctk.CTkFrame(title_bar, fg_color="transparent")
        theme_frame.pack(side="right")
        ctk.CTkLabel(theme_frame, text="Theme:", font=("Arial", 11)).pack(side="left", padx=(0, 5))
        self.theme_selector = ctk.CTkSegmentedButton(
            theme_frame,
            values=["System", "Hell", "Dunkel"],
            command=self.change_theme,
            width=200
        )
        self.theme_selector.set("System")
        self.theme_selector.pack(side="left")

        # ===== FILES SECTION =====
        files_frame = ctk.CTkFrame(main_container)
        files_frame.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(files_frame, text="Dateien", font=("Arial", 14, "bold")).pack(anchor="w", padx=15, pady=(10, 5))

        # Excel file
        excel_row = ctk.CTkFrame(files_frame, fg_color="transparent")
        excel_row.pack(fill="x", padx=15, pady=3)
        ctk.CTkLabel(excel_row, text="Excel-Datei:", anchor="w", width=100).pack(side="left")
        self.excel_entry = ctk.CTkEntry(excel_row, placeholder_text="beschreibungen.xlsx")
        self.excel_entry.pack(side="left", fill="x", expand=True, padx=10)
        self.excel_entry.bind("<FocusOut>", lambda e: self.save_current_settings())
        ctk.CTkButton(excel_row, text="...", width=40, command=self.browse_excel).pack(side="right")

        # Image folder
        folder_row = ctk.CTkFrame(files_frame, fg_color="transparent")
        folder_row.pack(fill="x", padx=15, pady=3)
        ctk.CTkLabel(folder_row, text="Bilder-Ordner:", anchor="w", width=100).pack(side="left")
        self.folder_entry = ctk.CTkEntry(folder_row, placeholder_text="pics/")
        self.folder_entry.pack(side="left", fill="x", expand=True, padx=10)
        self.folder_entry.bind("<FocusOut>", lambda e: self.save_current_settings())
        ctk.CTkButton(folder_row, text="...", width=40, command=self.browse_folder).pack(side="right")

        # Output file
        output_row = ctk.CTkFrame(files_frame, fg_color="transparent")
        output_row.pack(fill="x", padx=15, pady=(3, 10))
        ctk.CTkLabel(output_row, text="Ausgabedatei:", anchor="w", width=100).pack(side="left")
        self.output_entry = ctk.CTkEntry(output_row, placeholder_text="output_document.docx")
        self.output_entry.pack(side="left", fill="x", expand=True, padx=10)
        self.output_entry.bind("<FocusOut>", lambda e: self.save_current_settings())
        ctk.CTkButton(output_row, text="...", width=40, command=self.browse_output).pack(side="right")

        # ===== CAPTION COLUMNS SECTION =====
        caption_frame = ctk.CTkFrame(main_container)
        caption_frame.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(caption_frame, text="Bildunterschriften", font=("Arial", 14, "bold")).pack(anchor="w", padx=15, pady=(10, 5))

        # Column selection
        col_row = ctk.CTkFrame(caption_frame, fg_color="transparent")
        col_row.pack(fill="x", padx=15, pady=3)
        ctk.CTkLabel(col_row, text="Spalten:", anchor="w", width=100).pack(side="left")
        self.caption_cols_entry = ctk.CTkEntry(col_row, placeholder_text="I oder A,B,I")
        self.caption_cols_entry.pack(side="left", fill="x", expand=True, padx=10)
        self.caption_cols_entry.bind("<FocusOut>", lambda e: self.save_current_settings())

        # Separator
        sep_row = ctk.CTkFrame(caption_frame, fg_color="transparent")
        sep_row.pack(fill="x", padx=15, pady=(3, 10))
        ctk.CTkLabel(sep_row, text="Trenner:", anchor="w", width=100).pack(side="left")
        self.separator_entry = ctk.CTkEntry(sep_row, placeholder_text=" - ", width=80)
        self.separator_entry.pack(side="left", padx=10)
        self.separator_entry.bind("<FocusOut>", lambda e: self.save_current_settings())

        # ===== LAYOUT & FONT SECTION (COMBINED) =====
        settings_frame = ctk.CTkFrame(main_container)
        settings_frame.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(settings_frame, text="Layout & Schrift", font=("Arial", 14, "bold")).pack(anchor="w", padx=15, pady=(10, 5))

        # Images per page
        ipp_row = ctk.CTkFrame(settings_frame, fg_color="transparent")
        ipp_row.pack(fill="x", padx=15, pady=3)
        ctk.CTkLabel(ipp_row, text="Bilder/Seite:", anchor="w", width=100).pack(side="left")
        self.images_per_page = ctk.CTkComboBox(ipp_row, values=[str(i) for i in range(1, 11)], width=80, command=lambda _: self.save_current_settings())
        self.images_per_page.pack(side="left", padx=10)

        # Font family and size on same row
        ctk.CTkLabel(ipp_row, text="Schrift:", anchor="w").pack(side="left", padx=(20, 5))
        self.font_family = ctk.CTkComboBox(ipp_row, values=["Arial", "Times New Roman", "Calibri", "Helvetica"], width=140, command=lambda _: self.save_current_settings())
        self.font_family.pack(side="left", padx=5)
        self.font_size = ctk.CTkComboBox(ipp_row, values=[str(i) for i in range(8, 25)], width=60, command=lambda _: self.save_current_settings())
        self.font_size.pack(side="left", padx=5)

        # Font style
        style_row = ctk.CTkFrame(settings_frame, fg_color="transparent")
        style_row.pack(fill="x", padx=15, pady=(3, 10))
        ctk.CTkLabel(style_row, text="", width=100).pack(side="left")  # Spacer
        self.font_bold = ctk.CTkCheckBox(style_row, text="Fett", width=80, command=self.save_current_settings)
        self.font_bold.pack(side="left", padx=10)
        self.font_italic = ctk.CTkCheckBox(style_row, text="Kursiv", width=80, command=self.save_current_settings)
        self.font_italic.pack(side="left", padx=5)
        self.font_underline = ctk.CTkCheckBox(style_row, text="Unterstrichen", width=120, command=self.save_current_settings)
        self.font_underline.pack(side="left", padx=5)

        # ===== TEST MODE SECTION =====
        test_frame = ctk.CTkFrame(main_container)
        test_frame.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(test_frame, text="Test-Modus", font=("Arial", 14, "bold")).pack(anchor="w", padx=15, pady=(10, 5))

        test_row = ctk.CTkFrame(test_frame, fg_color="transparent")
        test_row.pack(fill="x", padx=15, pady=(3, 10))
        self.test_mode = ctk.CTkCheckBox(test_row, text="Aktiviert", command=self.toggle_test_mode, width=100)
        self.test_mode.pack(side="left")
        ctk.CTkLabel(test_row, text="Anzahl Bilder:").pack(side="left", padx=(10, 5))
        self.test_limit = ctk.CTkComboBox(test_row, values=[str(i) for i in [10, 20, 30, 50, 100]], width=80, state="disabled", command=lambda _: self.save_current_settings())
        self.test_limit.pack(side="left")

        # ===== PROGRESS SECTION =====
        self.progress_frame = ctk.CTkFrame(main_container)
        self.progress_frame.pack(fill="x", pady=(0, 10))

        self.status_label = ctk.CTkLabel(self.progress_frame, text="⏸ Bereit", font=("Arial", 13))
        self.status_label.pack(anchor="w", padx=15, pady=(10, 5))

        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.pack(fill="x", padx=15, pady=5)
        self.progress_bar.set(0)

        self.progress_text = ctk.CTkLabel(self.progress_frame, text="", font=("Arial", 11), text_color="gray")
        self.progress_text.pack(anchor="w", padx=15, pady=(0, 10))

        # ===== ERROR PANEL (hidden by default) =====
        self.error_frame = ctk.CTkFrame(main_container)
        # Don't pack yet - will be shown when errors occur

        error_header = ctk.CTkFrame(self.error_frame, fg_color="transparent")
        error_header.pack(fill="x", padx=15, pady=(10, 5))
        ctk.CTkLabel(error_header, text="⚠️ Fehler", font=("Arial", 13, "bold"), text_color="#e68a00").pack(side="left")
        self.error_count_label = ctk.CTkLabel(error_header, text="", font=("Arial", 11), text_color="gray")
        self.error_count_label.pack(side="left", padx=(10, 0))

        # Scrollable error text
        self.error_text = ctk.CTkTextbox(self.error_frame, height=100, font=("Arial", 10))
        self.error_text.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        # ===== ACTION BUTTON =====
        self.action_button = ctk.CTkButton(
            main_container,
            text="Dokument erstellen",
            command=self.action_button_clicked,
            height=45,
            font=("Arial", 16, "bold"),
            fg_color="#2fa572",
            hover_color="#258759"
        )
        self.action_button.pack(pady=(0, 10))

    def change_theme(self, value):
        """Change application theme"""
        theme_map = {
            "System": "system",
            "Hell": "light",
            "Dunkel": "dark"
        }
        ctk.set_appearance_mode(theme_map.get(value, "system"))

        # Save theme preference immediately
        self.save_current_settings()

    def toggle_test_mode(self):
        """Enable/disable test limit dropdown"""
        if self.test_mode.get():
            self.test_limit.configure(state="normal")
        else:
            self.test_limit.configure(state="disabled")
        self.save_current_settings()

    def browse_excel(self):
        """Open file dialog for Excel file"""
        filename = filedialog.askopenfilename(
            title="Excel-Datei auswählen",
            filetypes=[("Excel Dateien", "*.xlsx *.xls"), ("Alle Dateien", "*.*")]
        )
        if filename:
            self.excel_entry.delete(0, "end")
            self.excel_entry.insert(0, filename)
            self.save_current_settings()

    def browse_folder(self):
        """Open folder dialog for image folder"""
        folder = filedialog.askdirectory(title="Bilder-Ordner auswählen")
        if folder:
            self.folder_entry.delete(0, "end")
            self.folder_entry.insert(0, folder)
            self.save_current_settings()

    def browse_output(self):
        """Open save dialog for output file"""
        filename = filedialog.asksaveasfilename(
            title="Ausgabedatei speichern als",
            defaultextension=".docx",
            filetypes=[("Word Dokument", "*.docx"), ("Alle Dateien", "*.*")]
        )
        if filename:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, filename)
            self.save_current_settings()

    def load_saved_config(self):
        """Load saved configuration into GUI"""
        # Only insert if field is empty to avoid duplicates
        if not self.excel_entry.get():
            self.excel_entry.insert(0, self.config.get('excel_file', ''))
        if not self.folder_entry.get():
            self.folder_entry.insert(0, self.config.get('image_folder', ''))
        if not self.output_entry.get():
            self.output_entry.insert(0, self.config.get('output_file', ''))

        # Caption columns
        if not self.caption_cols_entry.get():
            caption_cols = self.config.get('caption_columns', ['I'])
            if isinstance(caption_cols, list):
                self.caption_cols_entry.insert(0, ','.join(caption_cols))
            else:
                self.caption_cols_entry.insert(0, caption_cols)

        if not self.separator_entry.get():
            self.separator_entry.insert(0, self.config.get('caption_separator', ' - '))

        # Layout
        self.images_per_page.set(str(self.config.get('images_per_page', 3)))

        # Font
        self.font_family.set(self.config.get('font_name', 'Arial'))
        self.font_size.set(str(self.config.get('font_size', 10)))

        if self.config.get('font_bold', False):
            self.font_bold.select()
        if self.config.get('font_italic', False):
            self.font_italic.select()
        if self.config.get('font_underline', False):
            self.font_underline.select()

        # Test mode
        if self.config.get('test_mode', False):
            self.test_mode.select()
            self.test_limit.configure(state="normal")
        self.test_limit.set(str(self.config.get('test_image_limit', 10)))

        # Theme (load saved theme) - don't trigger save
        saved_theme = self.config.get('theme', 'System')
        self.theme_selector.set(saved_theme)
        # Apply theme without saving
        theme_map = {
            "System": "system",
            "Hell": "light",
            "Dunkel": "dark"
        }
        ctk.set_appearance_mode(theme_map.get(saved_theme, "system"))

    def get_current_config(self):
        """Get configuration from GUI inputs"""
        # Parse caption columns
        caption_cols_str = self.caption_cols_entry.get().strip()
        caption_cols = [col.strip().upper() for col in caption_cols_str.split(',') if col.strip()]
        if not caption_cols:
            caption_cols = ['I']

        # Get test limit safely
        test_limit_value = 10  # default
        try:
            test_limit_value = int(self.test_limit.get())
        except (ValueError, AttributeError):
            pass  # Use default

        config = {
            'excel_file': self.excel_entry.get(),
            'image_folder': self.folder_entry.get(),
            'output_file': self.output_entry.get(),
            'filename_column': 'A',
            'caption_columns': caption_cols,
            'caption_separator': self.separator_entry.get() or ' - ',
            'images_per_page': int(self.images_per_page.get()),
            'font_name': self.font_family.get(),
            'font_size': int(self.font_size.get()),
            'font_bold': self.font_bold.get() == 1,
            'font_italic': self.font_italic.get() == 1,
            'font_underline': self.font_underline.get() == 1,
            'test_mode': self.test_mode.get() == 1,
            'test_image_limit': test_limit_value,
            'theme': self.theme_selector.get(),  # Save current theme
            'smart_layout': True,  # Always enabled
            'margin_top_cm': 1.27,    # Standard margins
            'margin_bottom_cm': 1.27,
            'margin_left_cm': 1.27,
            'margin_right_cm': 1.27,
        }
        return config

    def action_button_clicked(self):
        """Handle action button click (Start/Cancel)"""
        if self.is_processing:
            # Cancel processing
            self.cancel_processing = True
            self.update_status("⏹ Abbrechen...")
            self.action_button.configure(state="disabled")
        else:
            # Start processing
            self.start_processing()

    def start_processing(self):
        """Start document generation in background thread"""
        # Get configuration
        config = self.get_current_config()

        # Validate
        if not config['excel_file'] or not config['image_folder'] or not config['output_file']:
            self.status_label.configure(text="❌ Fehler: Bitte alle Dateien angeben!")
            return

        # Check if output file exists and warn
        output_path = Path(config['output_file'])
        if output_path.exists():
            result = messagebox.askyesno(
                "Datei überschreiben?",
                f"Die Datei '{output_path.name}' existiert bereits.\n\nMöchten Sie sie überschreiben?",
                icon='warning'
            )
            if not result:
                return

        # Save configuration
        self.config_manager.save_config(config)

        # Clear previous errors
        self.error_list = []
        self.error_frame.pack_forget()  # Hide error panel

        # Reset cancel flag
        self.cancel_processing = False

        # Update UI
        self.is_processing = True
        self.action_button.configure(
            text="Abbrechen",
            fg_color="#e63946",
            hover_color="#d62828"
        )
        self.status_label.configure(text="⏳ Verarbeitung läuft...")
        self.progress_bar.set(0)

        # Start processing thread
        self.processing_thread = threading.Thread(target=self.process_document, args=(config,))
        self.processing_thread.daemon = True
        self.processing_thread.start()

    def process_document(self, config):
        """Process document in background (runs in thread)"""
        try:
            # Read Excel
            if self.cancel_processing:
                return
            self.update_status("Lese Excel-Datei...")
            excel_reader = ExcelReader()
            excel_data = excel_reader.read_data(
                config['excel_file'],
                config['filename_column'],
                config['caption_columns'],
                config['caption_separator']
            )

            # Apply test limit
            if config['test_mode']:
                excel_data = excel_data[:config['test_image_limit']]

            # Process images
            if self.cancel_processing:
                return
            self.update_status("Suche Bilder...")
            image_handler = ImageHandler(config['image_folder'])
            complete_data = []

            for filename, caption in excel_data:
                if self.cancel_processing:
                    return
                try:
                    image_info = image_handler.get_image_info(filename)
                    complete_data.append((filename, caption, image_info['path'], image_info))
                except Exception as e:
                    self.error_list.append((filename, str(e)))
                    continue

            if not complete_data:
                self.update_status("❌ Keine Bilder gefunden!")
                self.processing_complete()
                return

            # Generate document
            if self.cancel_processing:
                return
            doc_generator = DocumentGenerator(config)
            processed, errors = doc_generator.create_document(
                complete_data,
                config['output_file'],
                progress_callback=self.update_progress_with_cancel_check
            )

            # Check if cancelled
            if self.cancel_processing:
                self.update_status("⏹ Abgebrochen")
                return

            # Done
            self.update_status(f"✓ Fertig! {processed}/{len(complete_data)} Bilder verarbeitet")
            self.progress_bar.set(1.0)

            # Show errors if any
            if self.error_list:
                self.show_errors()

        except Exception as e:
            if not self.cancel_processing:
                self.update_status(f"❌ Fehler: {str(e)}")
                # Show exception in error panel
                self.error_list.append(("SYSTEMFEHLER", str(e)))
                self.show_errors()
        finally:
            self.processing_complete()

    def update_progress_with_cancel_check(self, current, total, filename):
        """Update progress and check for cancellation"""
        if self.cancel_processing:
            # Signal to document generator to stop (would need implementation)
            return
        self.update_progress(current, total, filename)

    def update_status(self, text):
        """Update status label (thread-safe)"""
        self.after(0, lambda: self.status_label.configure(text=text))

    def update_progress(self, current, total, filename):
        """Update progress bar and text (thread-safe)"""
        progress = current / total if total > 0 else 0
        self.after(0, lambda: self.progress_bar.set(progress))
        self.after(0, lambda: self.progress_text.configure(
            text=f"Verarbeite: {filename} ({current}/{total})"
        ))

    def processing_complete(self):
        """Reset UI after processing completes"""
        self.is_processing = False
        self.cancel_processing = False
        self.after(0, lambda: self.action_button.configure(
            text="Dokument erstellen",
            fg_color="#2fa572",
            hover_color="#258759",
            state="normal"
        ))

    def show_errors(self):
        """Display error panel with all collected errors"""
        # Update count label
        count_text = f"({len(self.error_list)} Fehler)"
        self.after(0, lambda: self.error_count_label.configure(text=count_text))

        # Format error text
        error_text = "\n".join([f"❌ {filename}: {error}" for filename, error in self.error_list])

        # Update textbox and show panel
        def update_errors():
            self.error_text.delete("1.0", "end")
            self.error_text.insert("1.0", error_text)
            self.error_frame.pack(fill="x", pady=(0, 10), before=self.action_button)

        self.after(0, update_errors)

    def bring_to_foreground(self):
        """Bring application window to foreground (macOS specific)"""
        import platform
        if platform.system() == 'Darwin':  # macOS
            # Raise window and force focus
            self.lift()
            self.focus_force()
            self.attributes('-topmost', True)
            self.after(100, lambda: self.attributes('-topmost', False))

    def save_current_settings(self):
        """Save current GUI settings to configuration file"""
        # Don't save during initial load
        if self.is_loading:
            return

        try:
            config = self.get_current_config()
            self.config_manager.save_config(config)
        except Exception as e:
            print(f"Warning: Could not save configuration: {e}")

    def on_closing(self):
        """Handle window close event - save settings before closing"""
        self.save_current_settings()
        self.destroy()


def main():
    """Run the GUI"""
    app = Pic2DocGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
