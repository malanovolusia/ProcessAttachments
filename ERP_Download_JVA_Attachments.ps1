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

# Global variables
$script:JVAattachCount = 0
$script:JVAattachTotal = 0

#region Helper Functions

function Get-RandomHexValue {
    return [System.Guid]::NewGuid().ToString("N")
}

function Save-BinaryData {
    param([string]$FilePath, [object]$BinaryData)
    try {
        if ($null -eq $BinaryData) { return 1 }
        $directory = [System.IO.Path]::GetDirectoryName($FilePath)
        if (-not (Test-Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }
        [System.IO.File]::WriteAllBytes($FilePath, $BinaryData)
        return 0
    } catch {
        WriteLog "Error saving file: $_"
        return 1
    }
}

function Format-PadDate {
    param([object]$DateValue)
    if ($null -eq $DateValue -or $DateValue -eq [DBNull]::Value) { return "" }
    try {
        return ([DateTime]$DateValue).ToString("MM/dd/yyyy")
    } catch {
        return ""
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

function Get-JVAAttachments {
    param(
        [System.Data.Odbc.OdbcConnection]$ERPConnection,
        [System.Data.SqlClient.SqlConnection]$OnBaseConnection,
        [string]$DOC_CD,
        [string]$DOC_DEPT_CD,
        [string]$DOC_ID,
        [string]$DOC_VERS_NO,
        [string]$OBJ_ATT_PG_UNID,
        [string]$OutPath,
        [System.IO.StreamWriter]$IndexFileStream,
        [string]$CURR_FY,
        [string]$CURR_PER
    )

    $fileCount = 0

    # First query: Get metadata WITHOUT blob data
    $sqlMetadata = @"
SELECT a.OBJ_ATT_UNID, a.OBJ_ATT_NM, a.OBJ_ATT_SG_UNID, a.OBJ_ATT_DT, a.OBJ_ATT_USER_ID, a.OBJ_ATT_DSCR,
       a.OBJ_ATT_SEQ_NO, a.OBJ_ATT_ST, a.OBJ_ATT_TYP, a.OBJ_ATT_COMP_NM, a.OBJ_ATT_COMP_DESC, a.OBJ_ATT_DEL_USID, a.OBJ_ATT_DEL_DT
FROM O_FINPROD.IN_OBJ_ATT_CTLG a, O_FINPROD.IN_OBJ_ATT_DOC_REF c
WHERE a.OBJ_ATT_UNID = c.OBJ_ATT_UNID
  AND a.OBJ_ATT_PG_UNID = '$OBJ_ATT_PG_UNID' AND c.DOC_TYP = 'JV' AND c.DOC_CD = '$DOC_CD'
  AND c.DOC_ID = '$DOC_ID' AND c.DOC_VERS_NO <= $DOC_VERS_NO
ORDER BY a.OBJ_ATT_UNID
"@

    $cmd = New-Object System.Data.Odbc.OdbcCommand($sqlMetadata, $ERPConnection)
    $cmd.CommandTimeout = 120
    $reader = $cmd.ExecuteReader()

    while ($reader.Read()) {
        $script:JVAattachCount++
        $fileCount++

        $OBJ_ATT_UNID = $reader["OBJ_ATT_UNID"]
        $fileName = $reader["OBJ_ATT_NM"]
        
        # Check OnBase FIRST before pulling blob
        $existsInOnBase = Test-OnBaseAttachmentExists -Connection $OnBaseConnection -AttachmentID $OBJ_ATT_UNID

        if ($existsInOnBase) {
            WriteLog "JVA $DOC_DEPT_CD $DOC_ID [v$DOC_VERS_NO] Attachment already in OnBase: [$OBJ_ATT_UNID] $fileName"
            continue
        }

        # Only NOW retrieve the blob data for this specific attachment
        $sqlBlob = @"
SELECT b.OBJ_ATT_DATA
FROM O_FINPROD.IN_OBJ_ATT_STOR b
WHERE b.OBJ_ATT_UNID = '$OBJ_ATT_UNID'
"@

        $cmdBlob = New-Object System.Data.Odbc.OdbcCommand($sqlBlob, $ERPConnection)
        $cmdBlob.CommandTimeout = 120
        $readerBlob = $cmdBlob.ExecuteReader()

        if ($readerBlob.Read()) {
            $blobData = if ($readerBlob["OBJ_ATT_DATA"] -ne [DBNull]::Value) { 
                $readerBlob["OBJ_ATT_DATA"] 
            } else { 
                $null 
            }

            if ($null -ne $blobData) {
                $script:JVAattachTotal++
                
                # Read other metadata from outer reader
                $OBJ_ATT_SG_UNID = if ($reader["OBJ_ATT_SG_UNID"] -ne [DBNull]::Value) { $reader["OBJ_ATT_SG_UNID"] } else { "" }
                $OBJ_ATT_DT = $reader["OBJ_ATT_DT"]
                $OBJ_ATT_USER_ID = if ($reader["OBJ_ATT_USER_ID"] -ne [DBNull]::Value) { $reader["OBJ_ATT_USER_ID"] } else { "" }
                $OBJ_ATT_DSCR = if ($reader["OBJ_ATT_DSCR"] -ne [DBNull]::Value) { $reader["OBJ_ATT_DSCR"] } else { "" }
                $OBJ_ATT_SEQ_NO = if ($reader["OBJ_ATT_SEQ_NO"] -ne [DBNull]::Value) { $reader["OBJ_ATT_SEQ_NO"] } else { 0 }
                $OBJ_ATT_ST = if ($reader["OBJ_ATT_ST"] -ne [DBNull]::Value) { $reader["OBJ_ATT_ST"] } else { 0 }
                $OBJ_ATT_TYP = if ($reader["OBJ_ATT_TYP"] -ne [DBNull]::Value) { $reader["OBJ_ATT_TYP"] } else { 0 }
                $OBJ_ATT_COMP_NM = if ($reader["OBJ_ATT_COMP_NM"] -ne [DBNull]::Value) { $reader["OBJ_ATT_COMP_NM"] } else { "" }
                $OBJ_ATT_COMP_DESC = if ($reader["OBJ_ATT_COMP_DESC"] -ne [DBNull]::Value) { $reader["OBJ_ATT_COMP_DESC"] } else { "" }

                # Generate filename
                $extension = [System.IO.Path]::GetExtension($fileName)
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
                $guidHex = Get-RandomHexValue
                $fileNameGUID = if ($extension) { "${baseName}_[${guidHex}]${extension}" } else { "${fileName}_[${guidHex}]" }
                
                $fullPath = Join-Path $OutPath $fileNameGUID

                WriteLog "JVA $DOC_DEPT_CD $DOC_ID [v$DOC_VERS_NO] Saving attachment: [$fileCount] [$OBJ_ATT_UNID] $fileNameGUID"

                $result = Save-BinaryData -FilePath $fullPath -BinaryData $blobData

                if ($result -eq 0) {
                    $indexEntry = "$OBJ_ATT_UNID|$(Format-PadDate $OBJ_ATT_DT)|$fileNameGUID|$fullPath|$OBJ_ATT_USER_ID|$($OBJ_ATT_DSCR -replace "`r`n", " ")|$DOC_ID|$DOC_DEPT_CD|$DOC_ID|JV|$DOC_CD|$DOC_VERS_NO|$OBJ_ATT_SG_UNID|$OBJ_ATT_SEQ_NO|$OBJ_ATT_ST|$OBJ_ATT_TYP|$OBJ_ATT_COMP_NM|$OBJ_ATT_COMP_DESC|$fileName|$CURR_FY|$CURR_PER"
                    $IndexFileStream.WriteLine($indexEntry)
                }
            }
        }

        $readerBlob.Close()
    }

    $reader.Close()
}

#endregion

#region Main Processing

try {
    CreateGlobalVariables $PSCommandPath $PSScriptRoot
    
    $SID = "ERP19PRO"
    $OutputPath = "\\dlerp311birt\ProcessingCenter\Main\output\JETPDF\JVA\attachments"
    $CheckDuplicates = $true

    WriteLog("Log Files Named")

    WriteLog "Starting JVA Attachment Download"
    WriteLog "SID: $SID"
    WriteLog "Output Path: $OutputPath"

    # Ensure output directory exists
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    # Database connection strings
    $erpConnString = GetOracleConnectString $SID "PDI_USER" $False

    $onbaseServer = if ($SID -eq "ERP19PRO") { "obdbprod" } else { "obdbtest" }
    $onbaseDatabase = if ($SID -eq "ERP19PRO") { "OnBase" } else { "OB15TEST" }
    $onbaseConnString = "Server=$onbaseServer;Database=$onbaseDatabase;User Id=onbase_db_readonly;Password=p0exV3XanGknDfFnBvMe;"

    WriteLog "Connection String: $erpConnString"

    # Open database connections
    $erpConnection = New-Object System.Data.Odbc.OdbcConnection($erpConnString)
    $erpConnection.Open()
    WriteLog "Connected to ERP database"

    $onbaseConnection = New-Object System.Data.SqlClient.SqlConnection($onbaseConnString)
    $onbaseConnection.Open()
    WriteLog "Connected to OnBase database"

    # Open Index file (for tracking downloaded attachments)
    $indexFilePath = Join-Path $OutputPath "!JVA_attachment_index.txt"
    $indexFileStream = New-Object System.IO.StreamWriter($indexFilePath, $false, [System.Text.Encoding]::UTF8)

    # Write index header
    $indexFileStream.WriteLine("OBJ_ATT_UNID|ATT_DATE|GUID_FILENAME|FULL_PATH|USER_ID|DESCRIPTION|DOC_ID|DEPT_CD|DOC_ID_OUT|DOC_TYP|DOC_CD|VERS_NO|SG_UNID|SEQ_NO|STATUS|TYPE|COMP_NM|COMP_DESC|ORIGINAL_FILENAME|CURR_FY|CURR_PER")

    WriteLog "Index file: $indexFilePath"

    # Query for JVA documents with attachments
    $jvaQuery = @"
        SELECT DOC_CD, DOC_DEPT_CD, DOC_ID, DOC_VERS_NO, OBJ_ATT_PG_UNID, DOC_CREA_DT, CURR_FY, CURR_PER
        FROM O_FINPROD.JV_DOC_HDR
        WHERE OBJ_ATT_PG_UNID IS NOT NULL AND OBJ_ATT_PG_TOT > 0 AND DOC_CREA_DT > '05-OCT-2025' AND DOC_CD = 'JVA' AND DOC_PHASE_CD = 3
        ORDER BY DOC_ID, DOC_VERS_NO
        FETCH FIRST 10 ROW ONLY
"@


    WriteLog "Querying for JVA documents..."
    WriteLog "Main():Running SQL: $jvaQuery"

    $cmd = New-Object System.Data.Odbc.OdbcCommand($jvaQuery, $erpConnection)
    $cmd.CommandTimeout = 300
    $reader = $cmd.ExecuteReader()

    $docCount = 0

    while ($reader.Read()) {
        $docCount++

        $DOC_CD = $reader["DOC_CD"]
        $DOC_DEPT_CD = $reader["DOC_DEPT_CD"]
        $DOC_ID = $reader["DOC_ID"]
        $DOC_VERS_NO = $reader["DOC_VERS_NO"]
        $OBJ_ATT_PG_UNID = $reader["OBJ_ATT_PG_UNID"]
        $CURR_FY = if ($reader["CURR_FY"] -ne [DBNull]::Value) { $reader["CURR_FY"] } else { "" }
        $CURR_PER = if ($reader["CURR_PER"] -ne [DBNull]::Value) { $reader["CURR_PER"] } else { "" }

        WriteLog "Processing JVA #$docCount : $DOC_CD $DOC_DEPT_CD $DOC_ID v$DOC_VERS_NO (FY: $CURR_FY, PER: $CURR_PER)"

        # Get attachments for this JVA document
        Get-JVAAttachments `
            -ERPConnection $erpConnection `
            -OnBaseConnection $onbaseConnection `
            -DOC_CD $DOC_CD `
            -DOC_DEPT_CD $DOC_DEPT_CD `
            -DOC_ID $DOC_ID `
            -DOC_VERS_NO $DOC_VERS_NO `
            -OBJ_ATT_PG_UNID $OBJ_ATT_PG_UNID `
            -OutPath $OutputPath `
            -IndexFileStream $indexFileStream `
            -CURR_FY $CURR_FY `
            -CURR_PER $CURR_PER
    }

    $reader.Close()

    WriteLog ""
    WriteLog "========================================="
    WriteLog "Download Complete"
    WriteLog "Total JVA documents processed: $docCount"
    WriteLog "Total attachments found: $script:JVAattachCount"
    WriteLog "Total attachments downloaded: $script:JVAattachTotal"
    WriteLog "========================================="

    # Clean up
    $indexFileStream.Close()
    $erpConnection.Close()
    $onbaseConnection.Close()

    WriteLog "Attachment download completed successfully"

} catch {
    WriteLog "ERROR: $_"
    WriteLog $_.ScriptStackTrace
    exit 1
}

#endregion

