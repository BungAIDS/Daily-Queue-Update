# Launch the live queue watcher and KEEP THIS WINDOW OPEN afterward — so you can
# read the final output (and any error/traceback) even after you Ctrl+C it, the
# watch window closes at 5 PM, or it crashes. Run this instead of `python
# watch.py` for manual / debugging runs:
#
#     powershell -ExecutionPolicy Bypass -File .\watch.ps1
#   or, from an open PowerShell in the project folder:
#     .\watch.ps1
#
# All output is also saved to <project>\logs\watch.log (rotated daily, ~a week
# kept) regardless of this window, so nothing is lost.
#
# Ctrl+C once = stop cleanly (state is saved). Ctrl+C twice = force quit.
# Either way this window stays open until you press Enter.

$ErrorActionPreference = "Continue"
Set-Location -Path $PSScriptRoot

# Prefer the project venv's python; fall back to whatever 'python' is on PATH.
$py = Join-Path $PSScriptRoot "venv\Scripts\python.exe"
if (-not (Test-Path $py)) { $py = "python" }

try {
    & $py watch.py @args
}
finally {
    $code = $LASTEXITCODE
    Write-Host ""
    Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "Watcher exited (exit code $code). This window is staying open."   -ForegroundColor Yellow
    Write-Host "Full log: $(Join-Path $PSScriptRoot 'logs\watch.log')"            -ForegroundColor DarkGray
    Write-Host "Press Enter to close..."                                          -ForegroundColor Yellow
    [void](Read-Host)
}
