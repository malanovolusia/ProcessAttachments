#Requires -Version 5.1
<#
.SYNOPSIS
    Download JVA (Journal Voucher) attachments from ERP database
    
.EXAMPLE
    .\Download_JVA_Attachments.ps1
#>

Import-Module "\\erp311script\Library\PSM1\ERP_mod_logging.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_wmi.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_database.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_file.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_notify.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_datetime.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_util.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_string.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_print.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_hrm.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_fin.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_env.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_exec.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_interface.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_reporting.psm1"

$env:ENVIRONMENT = "PROD"

if ($env:ENVIRONMENT -eq "PROD") {
    # Production logging handled by module
} else {
    function WriteLog {
        param ([String]$message)
        Write-Host "Log: $message"
    }
}


#endregion

#region Main Processing

try {
    CreateGlobalVariables $PSCommandPath $PSScriptRoot

    WriteLog "========================================="
    WriteLog "Starting JVA Attachment Processing Orchestration"
    WriteLog "========================================="
    WriteLog ""

    # Define script paths
    $cleanupScript = Join-Path $PSScriptRoot "ERP_JVA_Attachments_CleanUp.ps1"
    $downloadScript = Join-Path $PSScriptRoot "ERP_JVA_Attachments_Downloader.ps1"
    $dipScript = Join-Path $PSScriptRoot "ERP_JVA_Attachments_DIP.ps1"
    $validationScript = Join-Path $PSScriptRoot "ERP_JVA_Attachments_Validation.ps1"
    $archiverScript = Join-Path $PSScriptRoot "ERP_JVA_Attachments_Archiver.ps1"

    # Verify scripts exist
    if (-not (Test-Path $cleanupScript)) {
        WriteLog "ERROR: Cleanup script not found at: $cleanupScript"
        exit 1
    }

    if (-not (Test-Path $downloadScript)) {
        WriteLog "ERROR: Download script not found at: $downloadScript"
        exit 1
    }

    if (-not (Test-Path $dipScript)) {
        WriteLog "ERROR: DIP script not found at: $dipScript"
        exit 1
    }

    if (-not (Test-Path $validationScript)) {
        WriteLog "ERROR: Validation script not found at: $validationScript"
        exit 1
    }

    if (-not (Test-Path $archiverScript)) {
        WriteLog "ERROR: Archiver script not found at: $archiverScript"
        exit 1
    }

    WriteLog "Scripts verified:"
    WriteLog "  - Cleanup: $cleanupScript"
    WriteLog "  - Download: $downloadScript"
    WriteLog "  - DIP: $dipScript"
    WriteLog "  - Validation: $validationScript"
    WriteLog "  - Archiver: $archiverScript"
    WriteLog ""

    # Step 0: Cleanup Old Files
    WriteLog "========================================="
    WriteLog "STEP 0: Cleaning Up Old Files"
    WriteLog "========================================="

    $cleanupStartTime = Get-Date

    try {
        $cleanupProcess = Start-Process -FilePath "powershell.exe" `
            -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$cleanupScript`"" `
            -Wait `
            -PassThru `
            -NoNewWindow

        if ($cleanupProcess.ExitCode -ne 0) {
            throw "Cleanup script failed with exit code: $($cleanupProcess.ExitCode)"
        }

        $cleanupEndTime = Get-Date
        $cleanupDuration = $cleanupEndTime - $cleanupStartTime

        WriteLog ""
        WriteLog "Cleanup completed successfully in $($cleanupDuration.TotalSeconds) seconds"
        WriteLog ""

    } catch {
        WriteLog "ERROR in cleanup step: $_"
        WriteLog $_.ScriptStackTrace
        exit 1
    }

    # Step 1: Download JVA Attachments
    WriteLog "========================================="
    WriteLog "STEP 1: Downloading JVA Attachments"
    WriteLog "========================================="

    $downloadStartTime = Get-Date

    try {
        $downloadProcess = Start-Process -FilePath "powershell.exe" `
            -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$downloadScript`"" `
            -Wait `
            -PassThru `
            -NoNewWindow

        if ($downloadProcess.ExitCode -ne 0) {
            throw "Download script failed with exit code: $($downloadProcess.ExitCode)"
        }

        $downloadEndTime = Get-Date
        $downloadDuration = $downloadEndTime - $downloadStartTime

        WriteLog ""
        WriteLog "Download completed successfully in $($downloadDuration.TotalSeconds) seconds"
        WriteLog ""

    } catch {
        WriteLog "ERROR in download step: $_"
        WriteLog $_.ScriptStackTrace
        exit 1
    }

    # Step 2: Create DIP Files
    WriteLog "========================================="
    WriteLog "STEP 2: Creating DIP Files"
    WriteLog "========================================="

    $dipStartTime = Get-Date

    try {
        $dipProcess = Start-Process -FilePath "powershell.exe" `
            -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$dipScript`"" `
            -Wait `
            -PassThru `
            -NoNewWindow

        if ($dipProcess.ExitCode -ne 0) {
            throw "DIP script failed with exit code: $($dipProcess.ExitCode)"
        }

        $dipEndTime = Get-Date
        $dipDuration = $dipEndTime - $dipStartTime

        WriteLog ""
        WriteLog "DIP file creation completed successfully in $($dipDuration.TotalSeconds) seconds"
        WriteLog ""

    } catch {
        WriteLog "ERROR in DIP creation step: $_"
        WriteLog $_.ScriptStackTrace
        exit 1
    }

    # Step 3: Wait before validation
    WriteLog "========================================="
    WriteLog "STEP 3: Waiting 15 minutes before validation"
    WriteLog "========================================="
    WriteLog "Waiting to allow OnBase to process the DIP file..."
    WriteLog "Wait started at: $(Get-Date -Format 'MM/dd/yyyy HH:mm:ss')"

    $waitSeconds = 600  # 15 minutes = 900 seconds
    Start-Sleep -Seconds $waitSeconds

    WriteLog "Wait completed at: $(Get-Date -Format 'MM/dd/yyyy HH:mm:ss')"
    WriteLog ""

    # Step 4: Validate Attachments in OnBase
    WriteLog "========================================="
    WriteLog "STEP 4: Validating Attachments in OnBase"
    WriteLog "========================================="

    $validationStartTime = Get-Date

    try {
        $validationProcess = Start-Process -FilePath "powershell.exe" `
            -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$validationScript`"" `
            -Wait `
            -PassThru `
            -NoNewWindow

        if ($validationProcess.ExitCode -ne 0) {
            throw "Validation script failed with exit code: $($validationProcess.ExitCode)"
        }

        $validationEndTime = Get-Date
        $validationDuration = $validationEndTime - $validationStartTime

        WriteLog ""
        WriteLog "Validation completed successfully in $($validationDuration.TotalSeconds) seconds"
        WriteLog ""

    } catch {
        WriteLog "ERROR in validation step: $_"
        WriteLog $_.ScriptStackTrace
        exit 1
    }

    # Step 5: Archive Processed Files
    WriteLog "========================================="
    WriteLog "STEP 5: Archiving Processed Files"
    WriteLog "========================================="

    $archiverStartTime = Get-Date

    try {
        $archiverProcess = Start-Process -FilePath "powershell.exe" `
            -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$archiverScript`"" `
            -Wait `
            -PassThru `
            -NoNewWindow

        if ($archiverProcess.ExitCode -ne 0) {
            throw "Archiver script failed with exit code: $($archiverProcess.ExitCode)"
        }

        $archiverEndTime = Get-Date
        $archiverDuration = $archiverEndTime - $archiverStartTime

        WriteLog ""
        WriteLog "Archiving completed successfully in $($archiverDuration.TotalSeconds) seconds"
        WriteLog ""

    } catch {
        WriteLog "ERROR in archiving step: $_"
        WriteLog $_.ScriptStackTrace
        exit 1
    }

    # Final Summary
    $totalEndTime = Get-Date
    $totalDuration = $totalEndTime - $cleanupStartTime

    WriteLog "========================================="
    WriteLog "JVA Attachment Processing Complete"
    WriteLog "========================================="
    WriteLog "Total execution time: $($totalDuration.TotalMinutes) minutes"
    WriteLog "  - Cleanup: $($cleanupDuration.TotalSeconds) seconds"
    WriteLog "  - Download: $($downloadDuration.TotalSeconds) seconds"
    WriteLog "  - DIP Creation: $($dipDuration.TotalSeconds) seconds"
    WriteLog "  - Wait Time: 600 seconds (10 minutes)"
    WriteLog "  - Validation: $($validationDuration.TotalSeconds) seconds"
    WriteLog "  - Archiving: $($archiverDuration.TotalSeconds) seconds"
    WriteLog "========================================="
    WriteLog ""
    WriteLog "Process completed successfully"

} catch {
    WriteLog "ERROR: $_"
    WriteLog $_.ScriptStackTrace
    exit 1
}

#endregion

