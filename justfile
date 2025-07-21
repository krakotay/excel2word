set shell := ["pwsh", "-Command"]

release:
    if (-not (Test-Path "release/backend")) { New-Item -ItemType Directory -Path "release/backend" }
    Get-ChildItem -Filter *.py -File | Copy-Item -Destination "release/backend" -Force
    Remove-Item -Path "release/*.7z" -Force -ErrorAction SilentlyContinue
    uv pip freeze > "requirements.txt"
    release/backend/python/python.exe -m pip install -r requirements.txt --break-system-packages --no-warn-script-location
    cd release && 7z a -mx=1 "EXCEL_В_WORD_0.5.7z" "."
