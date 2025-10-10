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

function Test-OnBaseAttachmentExists {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$AttachmentID
    )
    if (-not $CheckDuplicates) { return $false }
    try {
        $sql = "SELECT COUNT(*) AS NUM_FOUND FROM hsi.keyitem481 WHERE hsi.keyitem481.keyvaluebig = $AttachmentID AND (SELECT itemtypenum FROM hsi.itemdata WHERE hsi.itemdata.itemnum = hsi.keyitem481.itemnum) = 132"
        $cmd = New-Object System.Data.SqlClient.SqlCommand($sql, $Connection)
        $reader = $cmd.ExecuteReader()
        $count = 0
        if ($reader.Read()) { $count = $reader["NUM_FOUND"] }
        $reader.Close()
        return ($count -gt 0)
    } catch {
        WriteLog "Error checking OnBase: $_"
        return $false
    }
}


#region Main Processing

try {
    CreateGlobalVariables $PSCommandPath $PSScriptRoot
    
    $AttachmentsPath = "\\dlerp311birt\ProcessingCenter\Main\output\JETPDF\JVA\attachments"
    $IndexFilePath = Join-Path $AttachmentsPath "!JVA_attachment_index.txt"
    $DIPFilePath = Join-Path $AttachmentsPath "!JVA_attachment_indexes_DIP.txt"

    $SID = "ERP19PRO"
    
    # Database connection strings
    $erpConnString = GetOracleConnectString $SID "PDI_USER" $False

    $onbaseServer = if ($SID -eq "ERP19PRO") { "obdbprod" } else { "obdbtest" }
    $onbaseDatabase = if ($SID -eq "ERP19PRO") { "OnBase" } else { "OB15TEST" }
    $onbaseConnString = "Server=$onbaseServer;Database=$onbaseDatabase;User Id=onbase_db_readonly;Password=p0exV3XanGknDfFnBvMe;"

    $CheckDuplicates = $true

    # Check if index file exists
    if (-not (Test-Path $IndexFilePath)) {
        WriteLog "ERROR: Index file not found at: $IndexFilePath"
        exit 1
    }

    WriteLog "Starting OnBase validation"
    WriteLog "Index File: $IndexFilePath"
    WriteLog "OnBase Server: $onbaseServer"
    WriteLog "OnBase Database: $onbaseDatabase"

    # Connect to OnBase
    $onbaseConnection = New-Object System.Data.SqlClient.SqlConnection($onbaseConnString)
    $onbaseConnection.Open()
    WriteLog "Connected to OnBase database"

    # Read index file
    $indexLines = Get-Content $IndexFilePath
    if ($indexLines.Length -lt 2) {
        WriteLog "ERROR: Index file is empty or has no data rows"
        $onbaseConnection.Close()
        exit 1
    }

    $headerLine = $indexLines[0]
    $dataLines = $indexLines[1..($indexLines.Length - 1)]

    WriteLog "Found $($dataLines.Length) attachments to validate"

    # Validate each attachment
    $validatedCount = 0
    $foundCount = 0
    $notFoundCount = 0
    $errorCount = 0

    foreach ($line in $dataLines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $fields = $line -split '\|'
        if ($fields.Length -lt 1) { continue }

        $OBJ_ATT_UNID = $fields[0]
        $validatedCount++

        try {
            $existsInOnBase = Test-OnBaseAttachmentExists -Connection $onbaseConnection -AttachmentID $OBJ_ATT_UNID

            if ($existsInOnBase) {
                $foundCount++
                WriteLog "Found in OnBase: $OBJ_ATT_UNID"
            } else {
                $notFoundCount++
                WriteLog "NOT found in OnBase: $OBJ_ATT_UNID"
            }
        } catch {
            $errorCount++
            WriteLog "ERROR validating $OBJ_ATT_UNID : $_"
        }

        if ($validatedCount % 10 -eq 0) {
            WriteLog "Progress: Validated $validatedCount of $($dataLines.Length) attachments..."
        }
    }

    # Close connection
    $onbaseConnection.Close()

    # Summary
    WriteLog ""
    WriteLog "========================================="
    WriteLog "Validation Summary"
    WriteLog "========================================="
    WriteLog "Total attachments validated: $validatedCount"
    WriteLog "Found in OnBase: $foundCount"
    WriteLog "NOT found in OnBase: $notFoundCount"
    WriteLog "Errors: $errorCount"
    WriteLog "========================================="

    if ($notFoundCount -gt 0) {
        WriteLog "WARNING: $notFoundCount attachments were not found in OnBase"
    }

} catch {
    WriteLog "ERROR: $_"
    WriteLog $_.ScriptStackTrace
    exit 1
}

#endregion

