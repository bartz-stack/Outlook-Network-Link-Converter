@echo off
echo ========================================
echo Outlook Link Converter - Build Script
echo ========================================
echo.

REM Check if virtual environment exists
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    echo.
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat
echo.

REM Install/upgrade dependencies
echo Installing dependencies...
pip install --upgrade pip
pip install PyQt6
pip install pyinstaller
echo.

REM Clean previous builds
echo Cleaning previous builds...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del /q *.spec
echo.

REM Build executable with hidden imports
echo Building executable...
if exist "Outlook365_icon.ico" (
    echo Using icon file: Outlook365_icon.ico
    pyinstaller --onefile --windowed --icon=Outlook365_icon.ico --name "OutlookLinkConverter" --hidden-import=PyQt6.QtCore --hidden-import=PyQt6.QtGui --hidden-import=PyQt6.QtWidgets outlook_link_converter.py
) else (
    echo Icon file not found, building without icon
    pyinstaller --onefile --windowed --name "OutlookLinkConverter" --hidden-import=PyQt6.QtCore --hidden-import=PyQt6.QtGui --hidden-import=PyQt6.QtWidgets outlook_link_converter.py
)
echo.

REM Copy necessary files to dist folder
echo Copying files to dist folder...
copy gui_config.json dist\
if exist "Outlook365_icon.ico" copy Outlook365_icon.ico dist\
echo.

REM Create README for distribution
echo Creating distribution README...
(
echo Outlook Link Converter - Distribution Package
echo ============================================
echo.
echo FILES INCLUDED:
echo   - OutlookLinkConverter.exe   ^(Main application^)
echo   - gui_config.json            ^(Configuration file^)
echo   - Outlook365_icon.ico        ^(Icon file^)
echo.
echo INSTALLATION:
echo 1. Copy all files to any folder on your computer
echo 2. Double-click OutlookLinkConverter.exe to run
echo.
echo USAGE:
echo - Drag and drop files onto the window, OR
echo - Type/paste network paths manually
echo - Converted links are automatically copied to clipboard
echo - Paste into Outlook emails
echo.
echo CUSTOMIZATION:
echo - Edit gui_config.json to customize colors, text, and settings
echo - Changes take effect after restarting the application
echo.
echo NOTE: Keep all files together in the same folder!
) > dist\README.txt
echo.

echo ========================================
echo Build Complete!
echo ========================================
echo.
echo Files are located in the 'dist' folder:
dir /b dist
echo.
echo You can now distribute the entire 'dist' folder to users.
echo.
echo IMPORTANT - Windows Defender Warning:
echo If you get "Suspicious threat detected" this is a FALSE POSITIVE
echo This is extremely common with PyInstaller executables
echo.
echo To fix Windows Defender false positive:
echo 1. Open Windows Security
echo 2. Click "Virus ^& threat protection"
echo 3. Click "Protection history"
echo 4. Find "OutlookLinkConverter.exe"
echo 5. Click "Actions" -^> "Allow on device"
echo.
echo The .exe is safe - it's just your Python code packaged!
echo.
pause