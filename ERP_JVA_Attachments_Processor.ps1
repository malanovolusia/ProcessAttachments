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

$env:ENVIRONMENT = "DEV"

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
    $downloadScript = Join-Path $PSScriptRoot "ERP_Download_JVA_Attachments.ps1"
    $dipScript = Join-Path $PSScriptRoot "ERP_JVA_Attachments_DIP.ps1"

    # Verify scripts exist
    if (-not (Test-Path $downloadScript)) {
        WriteLog "ERROR: Download script not found at: $downloadScript"
        exit 1
    }

    if (-not (Test-Path $dipScript)) {
        WriteLog "ERROR: DIP script not found at: $dipScript"
        exit 1
    }

    WriteLog "Scripts verified:"
    WriteLog "  - Download: $downloadScript"
    WriteLog "  - DIP: $dipScript"
    WriteLog ""

    # Step 1: Download JVA Attachments
    WriteLog "========================================="
    WriteLog "STEP 1: Downloading JVA Attachments"
    WriteLog "========================================="

    $downloadStartTime = Get-Date

    try {
        & $downloadScript

        if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) {
            throw "Download script failed with exit code: $LASTEXITCODE"
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
        & $dipScript

        if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) {
            throw "DIP script failed with exit code: $LASTEXITCODE"
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

    # Final Summary
    $totalEndTime = Get-Date
    $totalDuration = $totalEndTime - $downloadStartTime

    WriteLog "========================================="
    WriteLog "JVA Attachment Processing Complete"
    WriteLog "========================================="
    WriteLog "Total execution time: $($totalDuration.TotalMinutes) minutes"
    WriteLog "  - Download: $($downloadDuration.TotalSeconds) seconds"
    WriteLog "  - DIP Creation: $($dipDuration.TotalSeconds) seconds"
    WriteLog "========================================="
    WriteLog ""
    WriteLog "Process completed successfully"

} catch {
    WriteLog "ERROR: $_"
    WriteLog $_.ScriptStackTrace
    exit 1
}

#endregion

