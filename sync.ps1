# sync.ps1 - Commit and push local changes to GitHub.
# Streamlit Cloud auto-redeploys after the push.
# Usage (double-click or via terminal):
#   powershell -ExecutionPolicy Bypass -File sync.ps1
#   powershell -ExecutionPolicy Bypass -File sync.ps1 -Message "your note"

param(
    [string]$Message = ""
)

$repoPath = "C:\Users\ruijie\.trae-cn\work\6a4f8645e4b01f7722a32ca7"
Set-Location -Path $repoPath

Write-Host "=== Git status ===" -ForegroundColor Cyan
git status --short

git add .

$changes = git diff --cached --name-only
if ($changes) {
    if ([string]::IsNullOrWhiteSpace($Message)) {
        $ts = Get-Date -Format "yyyyMMdd-HHmmss"
        $Message = "auto-sync $ts"
    }
    Write-Host "Committing with message: $Message" -ForegroundColor Yellow
    git commit -m $Message

    Write-Host "Pushing to origin main..." -ForegroundColor Green
    git push origin main

    Write-Host "Done. Streamlit Cloud redeploys in 1-3 minutes." -ForegroundColor Green
} else {
    Write-Host "No changes to commit. Already up to date." -ForegroundColor Green
}
