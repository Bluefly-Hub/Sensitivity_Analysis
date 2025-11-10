# Enable Windows Long Path Support (Run as Administrator)
# Use this before building .exe

New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem' -Name 'LongPathsEnabled' -Value 1 -PropertyType DWORD -Force -ErrorAction SilentlyContinue
Write-Host "`n✓ Long paths ENABLED" -ForegroundColor Green
Write-Host "You can now build the .exe with PyInstaller`n" -ForegroundColor Cyan
pause
