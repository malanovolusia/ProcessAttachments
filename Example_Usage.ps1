# Example Usage Script for JVA Attachment Processing
# This script demonstrates different ways to run the JVA attachment processor

# Example 1: Basic usage with defaults (TEST environment)
Write-Host "Example 1: Basic usage with defaults" -ForegroundColor Cyan
Write-Host ".\ERP_Process_JVA_Dip_Files.ps1" -ForegroundColor Yellow
Write-Host ""

# Example 2: Production environment with specific output path
Write-Host "Example 2: Production environment" -ForegroundColor Cyan
Write-Host '.\ERP_Process_JVA_Dip_Files.ps1 -SID "ERP19PRO" -OutputPath "\\erp311script\ProcessingCenter\Main\output\JETPDF\JVA\attachments"' -ForegroundColor Yellow
Write-Host ""

# Example 3: Test environment without duplicate checking
Write-Host "Example 3: Test without duplicate checking" -ForegroundColor Cyan
Write-Host '.\ERP_Process_JVA_Dip_Files.ps1 -SID "ERP19TEST" -OutputPath "C:\Temp\JVA_Test" -CheckDuplicates $false' -ForegroundColor Yellow
Write-Host ""

# Example 4: Local testing with custom path
Write-Host "Example 4: Local testing" -ForegroundColor Cyan
Write-Host '.\ERP_Process_JVA_Dip_Files.ps1 -OutputPath ".\Output\JVA\attachments"' -ForegroundColor Yellow
Write-Host ""

Write-Host "========================================" -ForegroundColor Green
Write-Host "To run one of these examples, copy and paste the command above" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# Prompt user to run
$response = Read-Host "Would you like to run Example 1 now? (Y/N)"
if ($response -eq 'Y' -or $response -eq 'y') {
    Write-Host "Running Example 1..." -ForegroundColor Green
    .\ERP_Process_JVA_Dip_Files.ps1
} else {
    Write-Host "No problem! You can run any example manually." -ForegroundColor Yellow
}

