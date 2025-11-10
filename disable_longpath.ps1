# Disable Windows Long Path Support (Run as Administrator)
# Use this after building .exe to restore normal file operation speed

Remove-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem' -Name 'LongPathsEnabled' -Force -ErrorAction SilentlyContinue
Write-Host "`n✓ Long paths DISABLED" -ForegroundColor Yellow
Write-Host "File operations should be faster now`n" -ForegroundColor Cyan
pause
