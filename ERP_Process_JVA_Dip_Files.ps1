#Requires -Version 5.1
<#
.SYNOPSIS
    Process JVA (Journal Voucher) attachments and create DIP files for OnBase import

.DESCRIPTION
    This script retrieves attachments for JVA documents from the ERP database,
    saves them to disk, and creates DIP (Document Import Package) files for OnBase.
    Based on the PO/DO/MA attachment processing logic from ERP_ncp_purchasing_post_processor.wsf

.PARAMETER SID
    Oracle SID (e.g., ERP19PRO, ERP19TEST)

.PARAMETER OutputPath
    Path where attachments and DIP files will be saved

.PARAMETER CheckDuplicates
    Whether to check if attachments already exist in OnBase (default: $true)

.EXAMPLE
    .\ERP_Process_JVA_Dip_Files.ps1 -SID "ERP19PRO" -OutputPath "\\server\share\JVA\attachments"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$SID = "ERP19TEST",

    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\Output\JVA\attachments",

    [Parameter(Mandatory=$false)]
    [bool]$CheckDuplicates = $true
)

# Import required modules
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

$env:ENVIRONMENT = "DEV"

$tmp = $ErrorActionPreference
$ErrorActionPreference = "Stop"

# Global variables
$script:JVAattachCount = 0
$script:JVAattachTotal = 0
$script:ProcessDate = Get-Date -Format "MM/dd/yyyy"

# Override WriteLog for DEV environment
if ($env:ENVIRONMENT -eq "PROD") {
    # Production logging handled by module
} else {
    function WriteLog {
        param ([String]$message)
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "[$timestamp] $message"
    }
}

try {
    CreateGlobalVariables $PSCommandPath $PSScriptRoot
    WriteLog "Log Files Named"
    WriteLog "Starting JVA Attachment Processing"
    WriteLog "SID: $SID"
    WriteLog "Output Path: $OutputPath"
    WriteLog "Check Duplicates: $CheckDuplicates"

} catch {
    WriteLog ($_ | Out-String).Trim()
    exit 1
} finally {
    $ErrorActionPreference = $tmp
}

#region Helper Functions

function Get-RandomHexValue {
    <#
    .SYNOPSIS
        Generate a random GUID-like hex value for unique filenames
    #>
    return [System.Guid]::NewGuid().ToString("N")
}

function Get-FileTypeNum {
    <#
    .SYNOPSIS
        Get OnBase file type number based on file extension
    #>
    param([string]$FileName)

    $extension = [System.IO.Path]::GetExtension($FileName).ToLower()

    switch ($extension) {
        ".pdf"  { return 16 }
        ".doc"  { return 17 }
        ".docx" { return 17 }
        ".xls"  { return 18 }
        ".xlsx" { return 18 }
        ".txt"  { return 19 }
        ".jpg"  { return 20 }
        ".jpeg" { return 20 }
        ".png"  { return 20 }
        ".tif"  { return 20 }
        ".tiff" { return 20 }
        default { return 16 }
    }
}

function Get-SHA256Hash {
    <#
    .SYNOPSIS
        Calculate SHA256 hash of a file
    #>
    param([string]$FilePath)

    if (-not (Test-Path $FilePath)) {
        return ""
    }

    try {
        $hash = Get-FileHash -Path $FilePath -Algorithm SHA256
        return $hash.Hash
    } catch {
        WriteLog "Error calculating SHA256 hash for $FilePath : $_"
        return ""
    }
}

function Save-BinaryData {
    <#
    .SYNOPSIS
        Save binary data (BLOB) to a file
    #>
    param(
        [string]$FilePath,
        [object]$BinaryData
    )

    try {
        if ($null -eq $BinaryData) {
            return 1
        }

        # Ensure directory exists
        $directory = [System.IO.Path]::GetDirectoryName($FilePath)
        if (-not (Test-Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }

        # Write binary data to file
        [System.IO.File]::WriteAllBytes($FilePath, $BinaryData)
        return 0
    } catch {
        WriteLog "Error saving binary data to $FilePath : $_"
        return 1
    }
}

function Format-PadDate {
    <#
    .SYNOPSIS
        Format date for DIP file (MM/DD/YYYY)
    #>
    param([object]$DateValue)

    if ($null -eq $DateValue -or $DateValue -eq [DBNull]::Value) {
        return ""
    }

    try {
        $date = [DateTime]$DateValue
        return $date.ToString("MM/dd/yyyy")
    } catch {
        return ""
    }
}

function Write-DIPHeader {
    <#
    .SYNOPSIS
        Write DIP file header
    #>
    param([System.IO.StreamWriter]$FileStream)

    # DIP files typically don't have a header, just documents
    # This is a placeholder if needed
}

function Write-DIPFooter {
    <#
    .SYNOPSIS
        Write DIP file footer with document count
    #>
    param(
        [int]$DocumentCount,
        [System.IO.StreamWriter]$FileStream
    )

    $FileStream.Write("END:")
    WriteLog "Total documents in DIP: $DocumentCount"
}

function Write-IndexHeader {
    <#
    .SYNOPSIS
        Write index file header
    #>
    param([System.IO.StreamWriter]$FileStream)

    $header = "OBJ_ATT_UNID|ATT_DATE|GUID_FILENAME|FULL_PATH|USER_ID|DESCRIPTION|DOC_ID|DEPT_CD|DOC_ID_OUT|DOC_TYP|DOC_CD|VERS_NO|SHA256"
    $FileStream.WriteLine($header)
}

function Write-IndexFooter {
    <#
    .SYNOPSIS
        Write index file footer
    #>
    param([System.IO.StreamWriter]$FileStream)

    # Index files typically don't have a footer
}

#endregion

#region Database Functions

function Test-OnBaseAttachmentExists {
    <#
    .SYNOPSIS
        Check if attachment already exists in OnBase
    #>
    param(
        [System.Data.Odbc.OdbcConnection]$Connection,
        [string]$AttachmentID
    )

    if (-not $CheckDuplicates) {
        return $false
    }

    try {
        $sql = @"
SELECT COUNT(*) AS NUM_FOUND
FROM hsi.keyitem481
WHERE hsi.keyitem481.keyvaluebig = $AttachmentID
AND (SELECT itemtypenum FROM hsi.itemdata WHERE hsi.itemdata.itemnum = hsi.keyitem481.itemnum) = 267
"@

        $cmd = New-Object System.Data.Odbc.OdbcCommand($sql, $Connection)
        $reader = $cmd.ExecuteReader()

        $count = 0
        if ($reader.Read()) {
            $count = $reader["NUM_FOUND"]
        }
        $reader.Close()

        return ($count -gt 0)
    } catch {
        WriteLog "Error checking OnBase for attachment $AttachmentID : $_"
        return $false
    }
}