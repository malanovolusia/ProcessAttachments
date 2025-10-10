#Requires -Version 5.1
<#
.SYNOPSIS
    Clean up JVA (Journal Voucher) attachment files

.DESCRIPTION
    Deletes all files from the JVA attachments folder.
    Logs all operations and provides summary of deleted files.

.EXAMPLE
    .\ERP_JVA_Attachments_CleanUp.ps1
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

    $CleanUpPath = "\\dlerp311birt\ProcessingCenter\Main\output\JETPDF\JVA\attachments"

    WriteLog "Starting cleanup process for: $CleanUpPath"

    # Verify the path exists
    if (-not (Test-Path -Path $CleanUpPath)) {
        WriteLog "ERROR: Path does not exist: $CleanUpPath"
        exit 1
    }

    # Get all files to delete (only in root folder, not subfolders)
    $FilesToDelete = Get-ChildItem -Path $CleanUpPath -File

    $FileCount = ($FilesToDelete | Measure-Object).Count
    WriteLog "Found $FileCount file(s) to delete"

    if ($FileCount -eq 0) {
        WriteLog "No files to delete. Cleanup complete."
        exit 0
    }

    # Delete the files
    $DeletedCount = 0
    $ErrorCount = 0

    foreach ($File in $FilesToDelete) {
        try {
            WriteLog "Deleting: $($File.FullName)"
            Remove-Item -Path $File.FullName -Force
            $DeletedCount++
        } catch {
            WriteLog "ERROR deleting file $($File.FullName): $_"
            $ErrorCount++
        }
    }

    WriteLog "Cleanup complete. Deleted: $DeletedCount file(s), Errors: $ErrorCount"

    if ($ErrorCount -gt 0) {
        exit 1
    }

} catch {
    WriteLog "ERROR: $_"
    WriteLog $_.ScriptStackTrace
    exit 1
}

#endregion

