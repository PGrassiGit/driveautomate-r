$ErrorActionPreference = "Stop"

python -m pip install -r requirements-lock.txt
python scripts/check_no_secrets.py
$env:QT_QPA_PLATFORM = "offscreen"
python -m pytest -q
Remove-Item Env:QT_QPA_PLATFORM -ErrorAction SilentlyContinue
python -m PyInstaller --noconfirm --clean DriveAutomate.spec

Write-Host "Executável gerado em: dist\DriveAutomate.exe"
Get-FileHash -Algorithm SHA256 "dist\DriveAutomate.exe"
