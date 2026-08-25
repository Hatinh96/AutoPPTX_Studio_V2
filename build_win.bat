@echo off
echo Building AutoPPTX Studio V2 for Windows...

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

python -m PyInstaller --noconfirm AutoPPTX_Studio_V2.spec

echo.
echo Build complete! Output located in dist/AutoPPTX_Studio_V2.exe
pause
