@echo off
echo ========================================
echo Cleaning Up Build Artifacts
echo ========================================
echo.

echo Deleting venv folder...
if exist "venv" (
    rmdir /s /q venv
    echo   ✓ Deleted venv/ (~500 MB saved^)
) else (
    echo   - venv/ not found
)

echo Deleting build folder...
if exist "build" (
    rmdir /s /q build
    echo   ✓ Deleted build/ (~50 MB saved^)
) else (
    echo   - build/ not found
)

echo Deleting dist folder...
if exist "dist" (
    rmdir /s /q dist
    echo   ✓ Deleted dist/ (~30 MB saved^)
) else (
    echo   - dist/ not found
)

echo Deleting __pycache__...
if exist "__pycache__" (
    rmdir /s /q __pycache__
    echo   ✓ Deleted __pycache__/
) else (
    echo   - __pycache__/ not found
)

echo Deleting spec files...
if exist "*.spec" (
    del /q *.spec
    echo   ✓ Deleted *.spec files
) else (
    echo   - No spec files found
)

echo.
echo ========================================
echo Cleanup Complete!
echo ========================================
echo.
echo Kept (essential files):
echo   ✓ outlook_link_converter.py
echo   ✓ gui_config.json
echo   ✓ Outlook365_icon.ico
echo   ✓ build.bat
echo   ✓ requirements.txt
echo.
echo To rebuild the .exe: Run build.bat
echo.
pause