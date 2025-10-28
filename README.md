<img width="437" height="415" alt="image" src="https://github.com/user-attachments/assets/5df901db-09e6-4da3-96bb-b56733b5d96f" />


# Outlook Link Converter

A simple GUI application to convert network drive paths into Outlook 365-compatible links.

Windows 11 and Outlook 365 broke traditional network path linking. This tool converts paths like:

P:/Machine Shop/Documents/Thrive Lubrication Chart.docx

Into Outlook-compatible links:

https://outlook.office.com/local/path/file://"P:/Machine%20Shop/Documents/Thrive%20Lubrication%20Chart.docx"

## Features

- ✅ **Drag & Drop Support**: Simply drag files onto the application
- ✅ **Auto-Copy to Clipboard**: Converted links are automatically copied
- ✅ **Manual Input**: Type or paste paths directly
- ✅ **Always on Top**: Optional toggle to keep window above other applications
- ✅ **Easy Customization**: Edit GUI without repackaging (see `gui_config.json`)
- ✅ **User-Friendly Interface**: Clean, modern design
- ✅ **Standard Window Controls**: Minimize, maximize, and close buttons

## Installation

### Option 1: Using Python (Recommended for IT)

1. **Install Python** (if not already installed):
   - Download from https://www.python.org/downloads/
   - During installation, check "Add Python to PATH"

2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the Application**:
   ```bash
   python outlook_link_converter.py
   ```

### Option 2: Create Executable (For End Users)

1. **Install PyInstaller**:
   ```bash
   pip install pyinstaller
   ```

2. **Create Executable**:
   ```bash
   pyinstaller --onefile --windowed --name "OutlookLinkConverter" outlook_link_converter.py
   ```

3. **Distribute**:
   - Find the executable in the `dist` folder
   - Copy `gui_config.json` to the same folder as the .exe
   - Distribute both files to users

## Usage

### Method 1: Drag and Drop
1. Open the application
2. Drag a file from your network drive onto the drop zone
3. The converted link is automatically copied to your clipboard
4. Paste into Outlook email

### Method 2: Manual Entry
1. Open the application
2. Type or paste the network path into the input field
3. Click "Convert to Outlook Link" or press Enter
4. The converted link is automatically copied to your clipboard
5. Paste into Outlook email

## Customizing the GUI

You can customize the appearance and text without modifying the code:

1. Open `gui_config.json` in any text editor (Notepad, VSCode, etc.)
2. Edit any values:

```json
{
    "window_title": "Your Company Link Converter",
    "window_width": 800,
    "window_height": 600,
    "instructions_text": "Your custom instructions here",
    "input_placeholder": "Your custom placeholder",
    "convert_button_text": "Convert Link",
    "output_label_text": "Your output label",
    "clear_button_text": "Reset",
    "always_on_top_text": "Keep window on top",
    "always_on_top_default": true,
    "font_family": "Arial",
    "font_size": 11,
    "theme_color": "#0078D4",
    "success_color": "#107C10",
    "background_color": "#FFFFFF"
}
```

3. Save the file
4. Restart the application to see changes

### Customization Options:

- **window_title**: Application title bar text
- **window_width/height**: Window size in pixels
- **instructions_text**: Help text shown above input
- **input_placeholder**: Gray text in empty input field
- **convert_button_text**: Main button text
- **output_label_text**: Label above output field
- **clear_button_text**: Clear button text
- **always_on_top_text**: Text for the "always on top" checkbox
- **always_on_top_default**: Whether "always on top" is enabled by default (true/false)
- **font_family**: Font name (e.g., "Segoe UI", "Arial", "Calibri")
- **font_size**: Base font size in points
- **theme_color**: Primary color (hex code)
- **success_color**: Success message color (hex code)
- **background_color**: Window background color (hex code)

## Always on Top Feature

The application includes a **"Keep window on top"** checkbox that keeps the converter window above all other applications. This is perfect for:

- **Drag & Drop workflow**: Keep the converter visible while browsing File Explorer
- **Single monitor setups**: No need to constantly switch between windows
- **Quick conversions**: Convert multiple files without losing the window

The checkbox can be toggled on/off at any time during use. You can set the default state in `gui_config.json` with `"always_on_top_default": true` or `false`.

## Supported Path Formats

The application accepts paths in various formats:

- `P:/Machine Shop/Documents/file.docx`
- `P:\Machine Shop\Documents\file.docx`
- `\\server\share\folder\file.docx`
- `"P:/Folder with spaces/file.docx"`

All formats are automatically converted to the correct Outlook link format.

## Troubleshooting

### Application won't start
- Ensure Python is installed correctly
- Run: `pip install --upgrade PyQt6`

### Links don't work in Outlook
- Ensure you're using Outlook 365 (desktop or web)
- Try clicking the link vs. Ctrl+Click
- Check if your organization has disabled this feature

### Configuration changes don't appear
- Ensure `gui_config.json` is in the same folder as the executable
- Check for JSON syntax errors (use a JSON validator)
- Restart the application after making changes

## Technical Notes

- Spaces are converted to `%20`
- Backslashes are converted to forward slashes
- Special characters are URL-encoded
- The format follows Outlook 365's local file protocol

## System Requirements

- Windows 10/11
- Python 3.8 or higher (if running from source)
- Outlook 365

## Support

For issues or questions:
1. Check the Troubleshooting section
2. Verify `gui_config.json` is valid JSON
3. Ensure all files are in the same directory

## License

Free to use and modify for your organization.
