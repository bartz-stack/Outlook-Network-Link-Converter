import sys
import urllib.parse
from pathlib import Path
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                             QTextEdit, QFrame, QCheckBox)
from PyQt6.QtCore import Qt, QMimeData
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QFont, QIcon
import json


class OutlookLinkConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.load_config()
        self.init_ui()
        
    def load_config(self):
        """Load GUI configuration from JSON file"""
        try:
            with open('gui_config.json', 'r') as f:
                self.config = json.load(f)
        except FileNotFoundError:
            # Default configuration
            self.config = {
                "window_title": "Outlook Link Converter",
                "window_width": 400,
                "window_height": 300,
                "instructions_text": "Drag and drop a file or paste path below:",
                "input_placeholder": "e.g., P:/Machine Shop/Documents/file.docx",
                "convert_button_text": "Convert to Outlook Link",
                "output_label_text": "Converted Link:",
                "clear_button_text": "Clear",
                "always_on_top_text": "Keep window on top",
                "always_on_top_default": True,
                "drop_zone_text": "Drop File Here",
                "drop_zone_font_size": 11,
                "font_family": "Segoe UI",
                "font_size": 9,
                "theme_color": "#0078D4",
                "success_color": "#107C10",
                "background_color": "#262626",
                "text_color": "#FFFFFF",
                "input_background": "#333333",
                "input_text_color": "#FFFFFF",
                "output_background": "#333333",
                "output_text_color": "#FFFFFF",
                "unc_to_drive_mappings": {
                    "\\\\Schuette-DC1\\production": "P:",
                    "\\\\Schuette-DC1": "G:"
                }
            }
            # Save default config
            self.save_default_config()
    
    def save_default_config(self):
        """Save default configuration to JSON file"""
        with open('gui_config.json', 'w') as f:
            json.dump(self.config, indent=4, fp=f)
    
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle(self.config["window_title"])
        self.setGeometry(100, 100, self.config["window_width"], self.config["window_height"])
        
        # Set window icon if it exists
        try:
            icon_path = Path("Outlook365_icon.ico")
            if icon_path.exists():
                self.setWindowIcon(QIcon(str(icon_path)))
        except Exception:
            pass  # If icon loading fails, continue without it
        
        # Enable drag and drop
        self.setAcceptDrops(True)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(8)
        
        # Checkbox for always on top (top right corner, no title label)
        title_layout = QHBoxLayout()
        title_layout.addStretch()
        
        # Always on top checkbox (smaller font, right-aligned)
        self.always_on_top_checkbox = QCheckBox(self.config["always_on_top_text"])
        self.always_on_top_checkbox.setFont(QFont(self.config["font_family"], self.config["font_size"] - 2))
        self.always_on_top_checkbox.setChecked(self.config.get("always_on_top_default", True))
        self.always_on_top_checkbox.setStyleSheet(f"""
            QCheckBox {{
                color: {self.config['text_color']};
                padding: 2px;
            }}
            QCheckBox::indicator {{
                width: 14px;
                height: 14px;
            }}
            QCheckBox::indicator:checked {{
                background-color: {self.config['theme_color']};
                border: 2px solid {self.config['theme_color']};
                border-radius: 3px;
            }}
            QCheckBox::indicator:unchecked {{
                background-color: {self.config['input_background']};
                border: 2px solid #555555;
                border-radius: 3px;
            }}
        """)
        self.always_on_top_checkbox.stateChanged.connect(self.toggle_always_on_top)
        title_layout.addWidget(self.always_on_top_checkbox)
        
        main_layout.addLayout(title_layout)
        
        # Set initial always-on-top state
        if self.config.get("always_on_top_default", True):
            self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        
        # Drop zone frame
        self.drop_frame = QFrame()
        self.drop_frame.setFrameShape(QFrame.Shape.Box)
        self.drop_frame.setStyleSheet(f"""
            QFrame {{
                border: 2px dashed {self.config['theme_color']};
                border-radius: 5px;
                background-color: {self.config['input_background']};
                padding: 10px;
            }}
        """)
        drop_layout = QVBoxLayout()
        
        drop_label = QLabel(f"📁 {self.config.get('drop_zone_text', 'Drop File Here')}")
        drop_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        drop_font = QFont(self.config["font_family"], self.config.get("drop_zone_font_size", 11))
        drop_label.setFont(drop_font)
        drop_label.setStyleSheet(f"color: {self.config['text_color']}; border: none;")
        drop_layout.addWidget(drop_label)
        
        self.drop_frame.setLayout(drop_layout)
        main_layout.addWidget(self.drop_frame)
        
        # Instructions
        instructions = QLabel(self.config["instructions_text"])
        instructions.setWordWrap(True)
        instructions.setFont(QFont(self.config["font_family"], self.config["font_size"]))
        instructions.setStyleSheet(f"color: {self.config['text_color']};")
        main_layout.addWidget(instructions)
        
        # Input field
        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText(self.config["input_placeholder"])
        self.input_field.setFont(QFont(self.config["font_family"], self.config["font_size"]))
        self.input_field.setStyleSheet(f"""
            QLineEdit {{
                padding: 6px;
                border: 1px solid #555555;
                border-radius: 3px;
                background-color: {self.config['input_background']};
                color: {self.config['input_text_color']};
            }}
            QLineEdit:focus {{
                border: 1px solid {self.config['theme_color']};
            }}
        """)
        self.input_field.returnPressed.connect(self.convert_link)
        main_layout.addWidget(self.input_field)
        
        # Convert button
        button_layout = QHBoxLayout()
        self.convert_button = QPushButton(self.config["convert_button_text"])
        self.convert_button.setFont(QFont(self.config["font_family"], self.config["font_size"], QFont.Weight.Bold))
        self.convert_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.config['theme_color']};
                color: white;
                padding: 8px;
                border: none;
                border-radius: 3px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: #005A9E;
            }}
            QPushButton:pressed {{
                background-color: #004578;
            }}
        """)
        self.convert_button.clicked.connect(self.convert_link)
        button_layout.addWidget(self.convert_button)
        
        # Clear button
        self.clear_button = QPushButton(self.config["clear_button_text"])
        self.clear_button.setFont(QFont(self.config["font_family"], self.config["font_size"]))
        self.clear_button.setStyleSheet("""
            QPushButton {
                background-color: #555555;
                color: #FFFFFF;
                padding: 8px;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #666666;
            }
        """)
        self.clear_button.clicked.connect(self.clear_fields)
        button_layout.addWidget(self.clear_button)
        
        main_layout.addLayout(button_layout)
        
        # Output label with copy button on the right
        output_header_layout = QHBoxLayout()
        output_label = QLabel(self.config["output_label_text"])
        output_label.setFont(QFont(self.config["font_family"], self.config["font_size"], QFont.Weight.Bold))
        output_label.setStyleSheet(f"color: {self.config['text_color']};")
        output_header_layout.addWidget(output_label)
        output_header_layout.addStretch()
        
        # Copy button (aligned to the right)
        self.copy_button = QPushButton("📋 Copy")
        self.copy_button.setFont(QFont(self.config["font_family"], self.config["font_size"]))
        # Set a fixed width to match the layout
        self.copy_button.setMinimumWidth(80)
        self.copy_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.config['theme_color']};
                color: white;
                padding: 8px;
                border: none;
                border-radius: 3px;
            }}
            QPushButton:hover {{
                background-color: #005A9E;
            }}
            QPushButton:pressed {{
                background-color: #004578;
            }}
        """)
        self.copy_button.clicked.connect(self.copy_to_clipboard)
        output_header_layout.addWidget(self.copy_button)
        
        main_layout.addLayout(output_header_layout)
        
        # Output field
        self.output_field = QTextEdit()
        self.output_field.setReadOnly(True)
        self.output_field.setFont(QFont("Courier New", self.config["font_size"] - 1))
        self.output_field.setStyleSheet(f"""
            QTextEdit {{
                padding: 6px;
                border: 1px solid #555555;
                border-radius: 3px;
                background-color: {self.config['output_background']};
                color: {self.config['output_text_color']};
            }}
        """)
        self.output_field.setMaximumHeight(60)
        main_layout.addWidget(self.output_field)
        
        # Status label
        self.status_label = QLabel("")
        self.status_label.setFont(QFont(self.config["font_family"], self.config["font_size"] - 1))
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.status_label)
        
        central_widget.setLayout(main_layout)
        
        # Set window background
        self.setStyleSheet(f"QMainWindow {{ background-color: {self.config['background_color']}; }}")
        central_widget.setStyleSheet(f"background-color: {self.config['background_color']};")
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """Handle drag enter event"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.drop_frame.setStyleSheet(f"""
                QFrame {{
                    border: 3px solid {self.config['theme_color']};
                    border-radius: 5px;
                    background-color: #404040;
                    padding: 10px;
                }}
            """)
    
    def dragLeaveEvent(self, event):
        """Handle drag leave event"""
        self.drop_frame.setStyleSheet(f"""
            QFrame {{
                border: 2px dashed {self.config['theme_color']};
                border-radius: 5px;
                background-color: {self.config['input_background']};
                padding: 10px;
            }}
        """)
    
    def dropEvent(self, event: QDropEvent):
        """Handle drop event"""
        self.drop_frame.setStyleSheet(f"""
            QFrame {{
                border: 2px dashed {self.config['theme_color']};
                border-radius: 5px;
                background-color: {self.config['input_background']};
                padding: 10px;
            }}
        """)
        
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files:
            file_path = files[0]
            self.input_field.setText(file_path)
            self.convert_link()
    
    def convert_link(self):
        """Convert the network path to Outlook link format"""
        input_path = self.input_field.text().strip()
        
        if not input_path:
            self.show_status("Please enter a file path!", "error")
            return
        
        # Clean up the path
        # Remove quotes if present
        input_path = input_path.strip('"').strip("'")
        
        # Ensure all slashes are backslashes
        input_path = input_path.replace('/', '\\')
        
        # Convert UNC paths to drive letters if mapped
        unc_mappings = self.config.get("unc_to_drive_mappings", {})
        # Sort by length (longest first) to match most specific paths first
        sorted_mappings = sorted(unc_mappings.items(), key=lambda x: len(x[0]), reverse=True)
        
        for unc_path, drive_letter in sorted_mappings:
            # Normalize the UNC path
            unc_path_normalized = unc_path.replace('/', '\\')
            
            # Check if input path starts with this UNC path (case-insensitive)
            if input_path.upper().startswith(unc_path_normalized.upper()):
                # Replace UNC path with drive letter
                remainder = input_path[len(unc_path_normalized):].lstrip('\\')
                # Ensure drive letter ends with colon
                if not drive_letter.endswith(':'):
                    drive_letter = drive_letter + ':'
                input_path = drive_letter + '\\' + remainder
                break
        
        # Encode only spaces and special characters, but preserve backslashes and colons as literals
        # Use a custom encoding that only encodes spaces
        encoded_path = input_path.replace(' ', '%20')
        
        # Create the Outlook link
        outlook_link = f'https://outlook.office.com/local/path/file://{encoded_path}'
        
        # Set the output
        self.output_field.setPlainText(outlook_link)
        
        # Copy to clipboard
        clipboard = QApplication.clipboard()
        clipboard.setText(outlook_link)
        
        self.show_status("✓ Link converted and copied to clipboard!", "success")
        
        # Set the output
        self.output_field.setPlainText(outlook_link)
        
        # Copy to clipboard
        clipboard = QApplication.clipboard()
        clipboard.setText(outlook_link)
        
        self.show_status("✓ Link converted and auto-copied to clipboard!", "success")
    
    def copy_to_clipboard(self):
        """Copy the output link to clipboard"""
        output_text = self.output_field.toPlainText().strip()
        if output_text:
            clipboard = QApplication.clipboard()
            clipboard.setText(output_text)
            self.show_status("✓ Copied to clipboard!", "success")
        else:
            self.show_status("No link to copy!", "error")
    
    def clear_fields(self):
        """Clear input and output fields"""
        self.input_field.clear()
        self.output_field.clear()
        self.status_label.clear()
    
    def toggle_always_on_top(self, state):
        """Toggle the always-on-top window flag"""
        if state == Qt.CheckState.Checked.value:
            self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowStaysOnTopHint)
        self.show()  # Required to apply the new window flags
    
    def show_status(self, message, status_type="success"):
        """Show status message"""
        if status_type == "success":
            color = self.config["success_color"]
        else:
            color = "#D13438"
        
        self.status_label.setText(message)
        self.status_label.setStyleSheet(f"color: {color}; font-weight: bold;")


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    window = OutlookLinkConverter()
    window.show()
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()