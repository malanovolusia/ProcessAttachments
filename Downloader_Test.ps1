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

#region Blob Download Functions

<#
.SYNOPSIS
    Downloads Oracle BLOB attachments to disk using chunked streaming.

.DESCRIPTION
    Retrieves attachment metadata and BLOB data from Oracle ERP database tables,
    then downloads files using efficient chunked reading via ODBC GetBytes().

.PARAMETER Connection
    Open ODBC connection to Oracle database

.PARAMETER OutputPath
    Directory path where files will be saved

.PARAMETER DocType
    Document type filter (e.g., 'JV')

.PARAMETER DocCode
    Document code filter (e.g., 'JVA')

.PARAMETER DocId
    Document ID filter (e.g., 'CV12251005')

.PARAMETER DocVersionNo
    Maximum document version number (default: 1)

.PARAMETER ChunkSize
    Size of chunks for streaming download in bytes (default: 16000)

.EXAMPLE
    Download-OracleAttachments -Connection $erpConn -OutputPath "C:\Attachments" -DocType "JV" -DocCode "JVA" -DocId "CV12251005"

.OUTPUTS
    Returns hashtable with download statistics:
    @{
        TotalFiles = 5
        SuccessCount = 5
        FailedCount = 0
        TotalBytes = 12345678
    }
#>
function Download-OracleAttachments {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Data.Odbc.OdbcConnection]$Connection,

        [Parameter(Mandatory=$true)]
        [string]$OutputPath,

        [Parameter(Mandatory=$true)]
        [string]$DocType,

        [Parameter(Mandatory=$true)]
        [string]$DocCode,

        [Parameter(Mandatory=$true)]
        [string]$DocId,

        [Parameter(Mandatory=$false)]
        [int]$DocVersionNo = 1,

        [Parameter(Mandatory=$false)]
        [int]$ChunkSize = 16000
    )

    # Ensure output directory exists
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        WriteLog "Created output directory: $OutputPath"
    }

    # Statistics
    $stats = @{
        TotalFiles = 0
        SuccessCount = 0
        FailedCount = 0
        TotalBytes = 0
    }

    # Step 1: Get attachment metadata (fast query without BLOB data)
    $metadataQuery = @"
SELECT a.OBJ_ATT_UNID, a.OBJ_ATT_NM, a.OBJ_ATT_SEQ_NO, a.OBJ_ATT_TYP,
       DBMS_LOB.GETLENGTH(b.OBJ_ATT_DATA) as BLOB_SIZE
FROM O_FINPROD.IN_OBJ_ATT_CTLG a, O_FINPROD.IN_OBJ_ATT_STOR b, O_FINPROD.IN_OBJ_ATT_DOC_REF c
WHERE a.OBJ_ATT_UNID = b.OBJ_ATT_UNID(+) AND a.OBJ_ATT_UNID = c.OBJ_ATT_UNID
  AND c.DOC_TYP = '$DocType' AND c.DOC_CD = '$DocCode'
  AND c.DOC_ID = '$DocId' AND c.DOC_VERS_NO <= $DocVersionNo
  AND b.OBJ_ATT_DATA IS NOT NULL
ORDER BY a.OBJ_ATT_UNID
"@

    WriteLog "Fetching attachment metadata for $DocType/$DocCode/$DocId..."

    try {
        $metadataCmd = New-Object System.Data.Odbc.OdbcCommand($metadataQuery, $Connection)
        $metadataCmd.CommandTimeout = 60
        $metadataReader = $metadataCmd.ExecuteReader()

        $attachments = @()
        while ($metadataReader.Read()) {
            $attachments += [PSCustomObject]@{
                UNID = $metadataReader["OBJ_ATT_UNID"]
                FileName = $metadataReader["OBJ_ATT_NM"]
                SeqNo = $metadataReader["OBJ_ATT_SEQ_NO"]
                Type = $metadataReader["OBJ_ATT_TYP"]
                Size = if ($metadataReader["BLOB_SIZE"] -ne [DBNull]::Value) { $metadataReader["BLOB_SIZE"] } else { 0 }
            }
        }
        $metadataReader.Close()

        $stats.TotalFiles = $attachments.Count
        WriteLog "Found $($attachments.Count) attachments to download"

    } catch {
        WriteLog "ERROR fetching metadata: $_"
        return $stats
    }

    # Step 2: Download each attachment using chunked streaming
    foreach ($attachment in $attachments) {
        try {
            $objAttUnid = $attachment.UNID
            $fileName = $attachment.FileName
            $seqNo = $attachment.SeqNo
            $blobSize = $attachment.Size

            if ($blobSize -eq 0) {
                WriteLog "Skipping $fileName - empty blob"
                $stats.FailedCount++
                continue
            }

            WriteLog "Downloading: $fileName (UNID: $objAttUnid, Size: $([math]::Round($blobSize/1024, 2)) KB)"

            # Create unique filename to avoid overwrites
            $safeFileName = $fileName -replace '[\\/:*?"<>|]', '_'
            $outputFile = Join-Path $OutputPath "${objAttUnid}_${seqNo}_${safeFileName}"

            # Download using chunked streaming
            $downloadResult = Download-BlobChunked -Connection $Connection -UNID $objAttUnid -OutputFile $outputFile -ExpectedSize $blobSize -ChunkSize $ChunkSize

            if ($downloadResult.Success) {
                $stats.SuccessCount++
                $stats.TotalBytes += $downloadResult.BytesDownloaded
                WriteLog "  SUCCESS: Downloaded $([math]::Round($downloadResult.BytesDownloaded/1024, 2)) KB in $([math]::Round($downloadResult.ElapsedSeconds, 2))s ($([math]::Round($downloadResult.SpeedKBps, 2)) KB/s)"
            } else {
                $stats.FailedCount++
                WriteLog "  FAILED: $($downloadResult.Error)"
            }

        } catch {
            $stats.FailedCount++
            WriteLog "  ERROR processing attachment $($attachment.FileName): $_"
            WriteLog "  $($_.ScriptStackTrace)"
        }
    }

    WriteLog "Download complete. Success: $($stats.SuccessCount), Failed: $($stats.FailedCount), Total: $($stats.TotalFiles)"

    return $stats
}

<#
.SYNOPSIS
    Downloads a single BLOB from Oracle using chunked streaming.

.DESCRIPTION
    Uses ODBC DataReader.GetBytes() to stream BLOB data in chunks,
    avoiding memory issues and Oracle RAW/VARCHAR limits.

.PARAMETER Connection
    Open ODBC connection to Oracle database

.PARAMETER UNID
    OBJ_ATT_UNID value identifying the attachment

.PARAMETER OutputFile
    Full path where the file will be saved

.PARAMETER ExpectedSize
    Expected size in bytes (for progress indicator)

.PARAMETER ChunkSize
    Size of chunks for streaming in bytes (default: 16000)

.OUTPUTS
    Returns hashtable with download result:
    @{
        Success = $true/$false
        BytesDownloaded = 12345
        ElapsedSeconds = 1.23
        SpeedKBps = 123.45
        Error = "error message" (if failed)
    }
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

                    # Progress indicator
                    if ($ExpectedSize -gt 0) {
                        $percentComplete = [math]::Round(($result.BytesDownloaded / $ExpectedSize) * 100, 1)
                        Write-Host "`r  Progress: $percentComplete% ($([math]::Round($result.BytesDownloaded/1024, 2)) KB / $([math]::Round($ExpectedSize/1024, 2)) KB)" -NoNewline
                    }
                }

                Write-Host ""  # New line after progress

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

#endregion




try {
    CreateGlobalVariables $PSCommandPath $PSScriptRoot
    
    $SID = "ERP19PRO"
    $OutputPath = "N:\Projects\35442 - JVA DIP\ProcessAttachments\Attachments"

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

    # Use the reusable download function
    $stats = Download-OracleAttachments `
        -Connection $erpConnection `
        -OutputPath $OutputPath `
        -DocType "JV" `
        -DocCode "JVA" `
        -DocId "CV12251005" `
        -DocVersionNo 1 `
        -ChunkSize 16000

    WriteLog "Download Statistics:"
    WriteLog "  Total Files: $($stats.TotalFiles)"
    WriteLog "  Successful: $($stats.SuccessCount)"
    WriteLog "  Failed: $($stats.FailedCount)"
    WriteLog "  Total Bytes: $([math]::Round($stats.TotalBytes/1024/1024, 2)) MB"

    $erpConnection.Close()
    $onbaseConnection.Close()


} catch {
    WriteLog "ERROR: $_"
    WriteLog $_.ScriptStackTrace
    exit 1
}





