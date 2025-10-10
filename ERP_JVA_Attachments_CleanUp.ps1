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

    $CleanUpPath = "\\dlerp311birt\ProcessingCenter\Main\output\JETPDF\JVA\attachments"

    WriteLog "========================================="
    WriteLog "Starting Cleanup Process"
    WriteLog "========================================="
    WriteLog "Target Path: $CleanUpPath"
    WriteLog "Current User: $env:USERNAME"
    WriteLog "Computer Name: $env:COMPUTERNAME"
    WriteLog ""

    # Verify the path exists
    if (-not (Test-Path -Path $CleanUpPath)) {
        WriteLog "ERROR: Path does not exist: $CleanUpPath"
        exit 1
    }

    WriteLog "Path verified - exists and is accessible"

    # Get all files to delete (only in root folder, not subfolders)
    $FilesToDelete = Get-ChildItem -Path $CleanUpPath -File

    $FileCount = ($FilesToDelete | Measure-Object).Count
    WriteLog "Found $FileCount file(s) to delete"
    WriteLog ""

    if ($FileCount -eq 0) {
        WriteLog "No files to delete. Cleanup complete."
        exit 0
    }

    # List first few files for verification
    $previewCount = [Math]::Min(5, $FileCount)
    WriteLog "Preview of files to delete (first $previewCount):"
    for ($i = 0; $i -lt $previewCount; $i++) {
        WriteLog "  - $($FilesToDelete[$i].Name)"
    }
    if ($FileCount -gt 5) {
        WriteLog "  ... and $($FileCount - 5) more file(s)"
    }
    WriteLog ""

    # Delete the files
    $DeletedCount = 0
    $ErrorCount = 0
    $SkippedCount = 0

    foreach ($File in $FilesToDelete) {
        try {
            WriteLog "Attempting to delete: $($File.Name)"

            # Use -LiteralPath instead of -Path to handle special characters like brackets []
            # -LiteralPath treats the path as literal, not as a wildcard pattern
            # Don't pre-check existence as network shares can have stale cache
            Remove-Item -LiteralPath $File.FullName -Force -ErrorAction Stop

            WriteLog "Successfully deleted: $($File.Name)"
            $DeletedCount++

        } catch [System.Management.Automation.ItemNotFoundException] {
            # File doesn't exist (already deleted or stale cache)
            WriteLog "File not found (may have been already deleted): $($File.Name)"
            $SkippedCount++
        } catch [System.UnauthorizedAccessException] {
            # Permission denied
            WriteLog "ERROR: Access denied - $($File.Name)"
            WriteLog "  Check file permissions and ensure file is not in use"
            $ErrorCount++
        } catch [System.IO.IOException] {
            # File in use or other IO error
            WriteLog "ERROR: File in use or IO error - $($File.Name)"
            WriteLog "  Exception: $($_.Exception.Message)"
            $ErrorCount++
        } catch {
            # Other errors
            WriteLog "ERROR deleting file $($File.Name): $_"
            WriteLog "  Exception Type: $($_.Exception.GetType().FullName)"
            WriteLog "  Exception Message: $($_.Exception.Message)"
            $ErrorCount++
        }
    }

    WriteLog ""
    WriteLog "========================================="
    WriteLog "Cleanup Summary"
    WriteLog "========================================="
    WriteLog "Total files found: $FileCount"
    WriteLog "Successfully deleted: $DeletedCount"
    WriteLog "Errors: $ErrorCount"
    WriteLog "Skipped (not found): $SkippedCount"
    WriteLog "========================================="

    if ($ErrorCount -gt 0) {
        WriteLog "WARNING: Some files could not be deleted due to errors"
        exit 1
    }

    if ($DeletedCount -eq 0 -and $SkippedCount -eq 0 -and $FileCount -gt 0) {
        WriteLog "WARNING: No files were deleted or skipped even though $FileCount file(s) were found"
        exit 1
    }

    if ($DeletedCount -gt 0) {
        WriteLog "Cleanup completed successfully - $DeletedCount file(s) deleted"
    } elseif ($SkippedCount -gt 0) {
        WriteLog "Cleanup completed - all files were already deleted (skipped: $SkippedCount)"
    }

} catch {
    WriteLog "ERROR: $_"
    WriteLog $_.ScriptStackTrace
    exit 1
}

#endregion

