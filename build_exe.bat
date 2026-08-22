@echo off
setlocal
python -m pip install -r requirements-lock.txt || exit /b 1
python scripts\check_no_secrets.py || exit /b 1
set QT_QPA_PLATFORM=offscreen
python -m pytest -q || exit /b 1
set QT_QPA_PLATFORM=
python -m PyInstaller --noconfirm --clean DriveAutomate.spec || exit /b 1
echo Executavel gerado em: dist\DriveAutomate.exe
certutil -hashfile dist\DriveAutomate.exe SHA256
endlocal
