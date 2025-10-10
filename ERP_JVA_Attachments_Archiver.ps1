#Requires -Version 5.1
<#
.SYNOPSIS
    Archiver
    
.EXAMPLE
    .\ERP_JVA_Attachments_Archiver.ps1
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

# Function to get unique archive folder name with YYYYMMDD pattern
function Get-UniqueArchiveFolderName {
    param (
        [string]$ArchiveBasePath,
        [string]$DatePattern = (Get-Date -Format "yyyyMMdd")
    )

    $targetPath = Join-Path -Path $ArchiveBasePath -ChildPath $DatePattern

    # If the folder doesn't exist, return the base pattern
    if (-not (Test-Path -Path $targetPath)) {
        return $DatePattern
    }

    # If it exists, find the next available number
    $counter = 1
    do {
        $newFolderName = "${DatePattern}-${counter}"
        $targetPath = Join-Path -Path $ArchiveBasePath -ChildPath $newFolderName
        $counter++
    } while (Test-Path -Path $targetPath)

    return $newFolderName
}

# Function to archive folder
function Copy-FolderToArchive {
    param (
        [string]$SourcePath,
        [string]$ArchiveBasePath
    )

    # Ensure archive base path exists
    if (-not (Test-Path -Path $ArchiveBasePath)) {
        WriteLog "Creating archive base path: $ArchiveBasePath"
        New-Item -Path $ArchiveBasePath -ItemType Directory -Force | Out-Null
    }

    # Get unique folder name
    $archiveFolderName = Get-UniqueArchiveFolderName -ArchiveBasePath $ArchiveBasePath
    $destinationPath = Join-Path -Path $ArchiveBasePath -ChildPath $archiveFolderName

    WriteLog "Archiving folder to: $destinationPath"

    # Copy the folder
    Copy-Item -Path $SourcePath -Destination $destinationPath -Recurse -Force

    WriteLog "Successfully archived to: $archiveFolderName"

    return $destinationPath
}



#region Main Processing

try {
    CreateGlobalVariables $PSCommandPath $PSScriptRoot

    $FolderPath = "\\dlerp311birt\ProcessingCenter\Main\output\JETPDF\JVA\attachments"
    $ArchivePath = "\\dlerp311birt\ProcessingCenter\Main\output\JETPDF\Archive\JVA\Attachments"

    if (Test-Path -Path $FolderPath) {
        $archivedPath = Copy-FolderToArchive -SourcePath $FolderPath -ArchiveBasePath $ArchivePath
        WriteLog "Folder archived to: $archivedPath"
    } else {
        WriteLog "WARNING: Source folder not found: $FolderPath"
    }

} catch {
    WriteLog "ERROR: $_"
    WriteLog $_.ScriptStackTrace
    exit 1
}

#endregion

