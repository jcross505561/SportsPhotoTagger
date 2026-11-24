@echo off
powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Hidden -File "%~dp000_Admin\Scripts\SportsPhotoTaggerWindow.ps1" %*
exit /b %errorlevel%
