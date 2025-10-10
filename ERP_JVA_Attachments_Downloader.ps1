#Requires -Version 5.1
<#
.SYNOPSIS
    Download JVA (Journal Voucher) attachments from ERP database

.EXAMPLE
    .\ERP_JVA_Attachments_Downloader.ps1
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

# Global variables
$script:JVAattachCount = 0
$script:JVAattachTotal = 0
$script:JVAattachSkipped = 0
$script:JVAattachMissingBlob = 0

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

<#
.SYNOPSIS
    Downloads a single BLOB from Oracle using chunked streaming.

.DESCRIPTION
    Uses ODBC DataReader.GetBytes() to stream BLOB data in chunks,
    avoiding memory issues and Oracle RAW/VARCHAR limits.
#>
function Download-BlobChunked {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Data.Odbc.OdbcConnection]$Connection,

        [Parameter(Mandatory=$true)]
        [long]$UNID,

        [Parameter(Mandatory=$true)]
        [string]$OutputFile,

        [Parameter(Mandatory=$false)]
        [long]$ExpectedSize = 0,

        [Parameter(Mandatory=$false)]
        [int]$ChunkSize = 16000
    )

    $result = @{
        Success = $false
        BytesDownloaded = 0
        ElapsedSeconds = 0
        SpeedKBps = 0
        Error = $null
    }

    $fileStream = $null
    $reader = $null
    $startTime = Get-Date

    try {
        # Query to get the BLOB column
        $blobQuery = @"
SELECT b.OBJ_ATT_DATA
FROM O_FINPROD.IN_OBJ_ATT_STOR b
WHERE b.OBJ_ATT_UNID = $UNID
"@

        $blobCmd = New-Object System.Data.Odbc.OdbcCommand($blobQuery, $Connection)
        $blobCmd.CommandTimeout = 300
        $reader = $blobCmd.ExecuteReader([System.Data.CommandBehavior]::SequentialAccess)

        if ($reader.Read()) {
            # Check if the blob is not null
            if (-not $reader.IsDBNull(0)) {
                $fileStream = [System.IO.File]::Create($OutputFile)

                # Read the BLOB in chunks using GetBytes
                $buffer = New-Object byte[] $ChunkSize
                $fieldOffset = 0  # Position in the BLOB field

                while ($true) {
                    # GetBytes(ordinal, fieldOffset, buffer, bufferOffset, length)
                    $bytesRead = $reader.GetBytes(0, $fieldOffset, $buffer, 0, $ChunkSize)

                    if ($bytesRead -eq 0) {
                        break  # No more data
                    }

                    # Write the chunk to file
                    $fileStream.Write($buffer, 0, $bytesRead)
                    $result.BytesDownloaded += $bytesRead
                    $fieldOffset += $bytesRead
                }

                $result.Success = $true

            } else {
                $result.Error = "BLOB data is NULL"
            }
        } else {
            $result.Error = "No record found for UNID $UNID"
        }

    } catch {
        $result.Error = $_.Exception.Message
    } finally {
        if ($reader -ne $null) { $reader.Close() }
        if ($fileStream -ne $null) {
            $fileStream.Close()
            $fileStream.Dispose()
        }

        $elapsed = (Get-Date) - $startTime
        $result.ElapsedSeconds = $elapsed.TotalSeconds
        $result.SpeedKBps = if ($elapsed.TotalSeconds -gt 0) {
            [math]::Round(($result.BytesDownloaded / 1024) / $elapsed.TotalSeconds, 2)
        } else { 0 }
    }

    return $result
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
            $script:JVAattachSkipped++
            WriteLog "JVA $DOC_DEPT_CD $DOC_ID [v$DOC_VERS_NO] Attachment already in OnBase: [$OBJ_ATT_UNID] $fileName"
            continue
        }

        # Read metadata from outer reader
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

        WriteLog "JVA $DOC_DEPT_CD $DOC_ID [v$DOC_VERS_NO] Downloading attachment: [$fileCount] [$OBJ_ATT_UNID] $fileNameGUID"

        # Use FAST chunked download method
        $downloadResult = Download-BlobChunked -Connection $ERPConnection -UNID $OBJ_ATT_UNID -OutputFile $fullPath -ChunkSize 16000

        if ($downloadResult.Success) {
            $script:JVAattachTotal++

            $sizeKB = [math]::Round($downloadResult.BytesDownloaded / 1024, 2)
            $speedKBps = $downloadResult.SpeedKBps
            WriteLog "JVA $DOC_DEPT_CD $DOC_ID [v$DOC_VERS_NO] SUCCESS: Downloaded $sizeKB KB in $([math]::Round($downloadResult.ElapsedSeconds, 2))s ($speedKBps KB/s)"

            # Write to index file
            $indexEntry = "$OBJ_ATT_UNID|$(Format-PadDate $OBJ_ATT_DT)|$fileNameGUID|$fullPath|$OBJ_ATT_USER_ID|$($OBJ_ATT_DSCR -replace "`r`n", " ")|$DOC_ID|$DOC_DEPT_CD|$DOC_ID|JV|$DOC_CD|$DOC_VERS_NO|$OBJ_ATT_SG_UNID|$OBJ_ATT_SEQ_NO|$OBJ_ATT_ST|$OBJ_ATT_TYP|$OBJ_ATT_COMP_NM|$OBJ_ATT_COMP_DESC|$fileName|$CURR_FY|$CURR_PER"
            $IndexFileStream.WriteLine($indexEntry)
        } else {
            $script:JVAattachMissingBlob++
            WriteLog "JVA $DOC_DEPT_CD $DOC_ID [v$DOC_VERS_NO] ERROR: Failed to download attachment: [$OBJ_ATT_UNID] $fileName - $($downloadResult.Error)"
        }
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
        FETCH FIRST 30 ROW ONLY
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
    WriteLog "Total attachments skipped (already in OnBase): $script:JVAattachSkipped"
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

