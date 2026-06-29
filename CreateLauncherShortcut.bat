@echo off
setlocal

set "TARGET=%~dp0RunLauncher.vbs"
set "WORKDIR=%~dp0"
set "NAME=Daily Queue Launcher.lnk"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$desktop=[Environment]::GetFolderPath('Desktop'); $ws=New-Object -ComObject WScript.Shell; $s=$ws.CreateShortcut((Join-Path $desktop '%NAME%')); $s.TargetPath='%TARGET%'; $s.WorkingDirectory='%WORKDIR%'; $s.IconLocation='%SystemRoot%\System32\shell32.dll,167'; $s.Description='Daily Queue script launcher'; $s.Save()"

echo Created desktop shortcut: %USERPROFILE%\Desktop\Daily Queue Launcher.lnk
pause
