@echo off
echo Building ppt2a4.exe...
pyinstaller --onefile --windowed --name="ppt2a4" --icon=icon.ico --hidden-import=win32com --hidden-import=win32com.client --hidden-import=fitz --hidden-import=reportlab ppt2a4.py
echo Done. Executable is in 'dist' folder.
pause