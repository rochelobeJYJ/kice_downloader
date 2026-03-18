@echo off
echo ====================================================
echo KICE Downloader Distribution Auto-Build Script
echo ====================================================

echo [1/4] Building KICE_Downloader via PyInstaller...
C:\Users\user\anaconda3\Scripts\pyinstaller.exe --noconfirm --onedir --windowed --icon icon.ico --name KICE_Downloader --add-data "icon.ico;." --add-binary "C:\Users\user\anaconda3\Library\bin\tcl86t.dll;." --add-binary "C:\Users\user\anaconda3\Library\bin\tk86t.dll;." --add-data "C:\Users\user\anaconda3\Library\lib\tcl8.6;tcl8.6" --add-data "C:\Users\user\anaconda3\Library\lib\tk8.6;tk8.6" main.py

echo.
echo [2/4] Copying README.md...
copy README.md "dist\KICE_Downloader\" /Y

echo.
echo [3/4] Compressing to ZIP file... (Please wait)
if exist KICE_Downloader_Distribution.zip del /F /Q KICE_Downloader_Distribution.zip
powershell -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'dist\KICE_Downloader' -DestinationPath 'KICE_Downloader_Distribution.zip' -Force"

echo.
echo [4/4] Cleaning up temporary build files...
if exist KICE_Downloader.spec del /F /Q KICE_Downloader.spec
if exist build rmdir /S /Q build
if exist dist rmdir /S /Q dist

echo.
echo ====================================================
echo Build Complete! Check KICE_Downloader_Distribution.zip
echo ====================================================
pause
