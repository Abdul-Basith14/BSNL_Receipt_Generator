# Cleanup Script - Run this after closing all Excel files
# This removes old/duplicate files keeping only the final conversion script and output

Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "           CLEANUP: Removing Old Files" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan

# Files to remove
$filesToRemove = @(
    "Generated_Cash_Receipts.xlsx",
    "Formatted_Cash_Receipts.xlsx",
    "generate_cash_receipts.py"
)

Write-Host "`nRemoving old files..." -ForegroundColor Yellow

foreach ($file in $filesToRemove) {
    if (Test-Path $file) {
        try {
            Remove-Item $file -Force -ErrorAction Stop
            Write-Host "  [OK] Removed: $file" -ForegroundColor Green
        }
        catch {
            Write-Host "  [ERROR] Could not remove: $file (file may be open in Excel)" -ForegroundColor Red
        }
    }
}

Write-Host "`n==================================================================" -ForegroundColor Cyan
Write-Host "           FINAL FILES IN DIRECTORY:" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan

Get-ChildItem *.py, *.xlsx | Where-Object { $_.Name -notlike "~$*" } | ForEach-Object {
    Write-Host "  - $($_.Name)" -ForegroundColor White
}

Write-Host "`n==================================================================" -ForegroundColor Cyan
Write-Host "  CLEANUP COMPLETE!" -ForegroundColor Green
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "`nFinal Structure:" -ForegroundColor Yellow
Write-Host "  - Dec -25.xlsx                     (Source data file)" -ForegroundColor White
Write-Host "  - Final_Cash_Receipts.xlsx         (Generated output)" -ForegroundColor White
Write-Host "  - generate_cash_receipts_final.py  (Conversion script)" -ForegroundColor White
