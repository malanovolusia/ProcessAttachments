#Requires -Version 5.1
<#
.SYNOPSIS
    Process JVA (Journal Voucher) attachments and create DIP files for OnBase import

.EXAMPLE
    .\ERP_Process_JVA_Dip_Files.ps1"
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

# Global variables
$script:JVAattachCount = 0
$script:JVAattachTotal = 0
$script:ProcessDate = Get-Date -Format "MM/dd/yyyy"

#region Helper Functions

function Get-RandomHexValue {
    return [System.Guid]::NewGuid().ToString("N")
}

function Get-FileTypeNum {
    param([string]$FileName)
    $extension = [System.IO.Path]::GetExtension($FileName).ToLower()
    switch ($extension) {
        ".pdf"  { return 16 }
        ".doc"  { return 17 }
        ".docx" { return 17 }
        ".xls"  { return 18 }
        ".xlsx" { return 18 }
        default { return 16 }
    }
}

function Get-SHA256Hash {
    param([string]$FilePath)
    if (-not (Test-Path $FilePath)) { return "" }
    try {
        $hash = Get-FileHash -Path $FilePath -Algorithm SHA256
        return $hash.Hash
    } catch {
        WriteLog "Error calculating SHA256 hash: $_"
        return ""
    }
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
        $sql = "SELECT COUNT(*) AS NUM_FOUND FROM hsi.keyitem481 WHERE hsi.keyitem481.keyvaluebig = $AttachmentID AND (SELECT itemtypenum FROM hsi.itemdata WHERE hsi.itemdata.itemnum = hsi.keyitem481.itemnum) = 267"
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
        [System.IO.StreamWriter]$DIPFileStream,
        [System.IO.StreamWriter]$IndexFileStream
    )

    $fileCount = 0

    # SQL to retrieve attachments for JVA documents
    $sql = @"
SELECT a.OBJ_ATT_UNID, a.OBJ_ATT_NM, b.OBJ_ATT_DATA, a.OBJ_ATT_SG_UNID, a.OBJ_ATT_DT, a.OBJ_ATT_USER_ID, a.OBJ_ATT_DSCR,
       a.OBJ_ATT_SEQ_NO, a.OBJ_ATT_ST, a.OBJ_ATT_TYP, a.OBJ_ATT_COMP_NM, a.OBJ_ATT_COMP_DESC, a.OBJ_ATT_DEL_USID, a.OBJ_ATT_DEL_DT
FROM O_FINPROD.IN_OBJ_ATT_CTLG a, O_FINPROD.IN_OBJ_ATT_STOR b, O_FINPROD.IN_OBJ_ATT_DOC_REF c
WHERE a.OBJ_ATT_UNID = b.OBJ_ATT_UNID(+) AND a.OBJ_ATT_UNID = c.OBJ_ATT_UNID
  AND a.OBJ_ATT_PG_UNID = '$OBJ_ATT_PG_UNID' AND c.DOC_TYP = 'JV' AND c.DOC_CD = '$DOC_CD'
  AND c.DOC_ID = '$DOC_ID' AND c.DOC_VERS_NO <= $DOC_VERS_NO
ORDER BY b.OBJ_ATT_UNID
"@

    try {
        $cmd = New-Object System.Data.Odbc.OdbcCommand($sql, $ERPConnection)
        $cmd.CommandTimeout = 120
        $reader = $cmd.ExecuteReader()

        while ($reader.Read()) {
            $script:JVAattachCount++
            $fileCount++

            # Read attachment metadata
            $OBJ_ATT_UNID = $reader["OBJ_ATT_UNID"]
            $fileName = $reader["OBJ_ATT_NM"]
            $blobData = if ($reader["OBJ_ATT_DATA"] -ne [DBNull]::Value) { $reader["OBJ_ATT_DATA"] } else { $null }
            $OBJ_ATT_SG_UNID = if ($reader["OBJ_ATT_SG_UNID"] -ne [DBNull]::Value) { $reader["OBJ_ATT_SG_UNID"] } else { "" }
            $OBJ_ATT_DT = $reader["OBJ_ATT_DT"]
            $OBJ_ATT_USER_ID = if ($reader["OBJ_ATT_USER_ID"] -ne [DBNull]::Value) { $reader["OBJ_ATT_USER_ID"] } else { "" }
            $OBJ_ATT_DSCR = if ($reader["OBJ_ATT_DSCR"] -ne [DBNull]::Value) { $reader["OBJ_ATT_DSCR"] } else { "" }
            $OBJ_ATT_SEQ_NO = if ($reader["OBJ_ATT_SEQ_NO"] -ne [DBNull]::Value) { $reader["OBJ_ATT_SEQ_NO"] } else { 0 }
            $OBJ_ATT_ST = if ($reader["OBJ_ATT_ST"] -ne [DBNull]::Value) { $reader["OBJ_ATT_ST"] } else { 0 }
            $OBJ_ATT_TYP = if ($reader["OBJ_ATT_TYP"] -ne [DBNull]::Value) { $reader["OBJ_ATT_TYP"] } else { 0 }
            $OBJ_ATT_COMP_NM = if ($reader["OBJ_ATT_COMP_NM"] -ne [DBNull]::Value) { $reader["OBJ_ATT_COMP_NM"] } else { "" }
            $OBJ_ATT_COMP_DESC = if ($reader["OBJ_ATT_COMP_DESC"] -ne [DBNull]::Value) { $reader["OBJ_ATT_COMP_DESC"] } else { "" }
            $OBJ_ATT_DEL_USID = if ($reader["OBJ_ATT_DEL_USID"] -ne [DBNull]::Value) { $reader["OBJ_ATT_DEL_USID"] } else { "" }
            $OBJ_ATT_DEL_DT = if ($reader["OBJ_ATT_DEL_DT"] -ne [DBNull]::Value) { $reader["OBJ_ATT_DEL_DT"] } else { "" }

            # Generate unique filename with GUID
            $extension = [System.IO.Path]::GetExtension($fileName)
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
            $guidHex = Get-RandomHexValue
            $fileNameGUID = if ($extension) { "${baseName}_[${guidHex}]${extension}" } else { "${fileName}_[${guidHex}]" }

            # Check if attachment already exists in OnBase
            $existsInOnBase = Test-OnBaseAttachmentExists -Connection $OnBaseConnection -AttachmentID $OBJ_ATT_UNID

            if (-not $existsInOnBase -and $null -ne $blobData) {
                $script:JVAattachTotal++
                $fullPath = Join-Path $OutPath $fileNameGUID

                WriteLog "JVA $DOC_DEPT_CD $DOC_ID [v$DOC_VERS_NO] Saving attachment: [$fileCount] [$OBJ_ATT_UNID] $fileNameGUID"

                # Save attachment to disk
                $result = Save-BinaryData -FilePath $fullPath -BinaryData $blobData

                if ($result -eq 0) {
                    $sha256Hash = Get-SHA256Hash -FilePath $fullPath
                    $descriptionField = if ($OBJ_ATT_DSCR.Trim().Length -gt 0) { "Long Description: $OBJ_ATT_DSCR`r`n" } else { "" }

                    # Build DIP entry
                    $dipEntry = @"
BEGIN:
>>Dummy Key: Document #$fileCount
>>DocTypeName: FIN - JVA Attachments
>>DocDate: $script:ProcessDate
Journal Voucher #: $DOC_ID
Advantage Attachment ID: $OBJ_ATT_UNID
Attachment Date: $(Format-PadDate $OBJ_ATT_DT)
${descriptionField}Filename: $fileName
>>Dummy Key: hidden keywords begin here
Doc ID: $DOC_ID
Version #: $DOC_VERS_NO
Department #: $DOC_DEPT_CD
Advantage Doc Type: JV
Advantage Doc Code: $DOC_CD
GUID File Name: $fileNameGUID
Advantage Attachment Primary Group ID: $OBJ_ATT_PG_UNID
Advantage Attachment Secondary Group ID: $OBJ_ATT_SG_UNID
Advantage Attachment User: $OBJ_ATT_USER_ID
Advantage Attachment Status: $OBJ_ATT_ST
Advantage Attachment Type: $OBJ_ATT_TYP
Advantage Attachment Component Name: $OBJ_ATT_COMP_NM
Advantage Attachment Component Context: $OBJ_ATT_COMP_DESC
Advantage Attachment Sequence #: $OBJ_ATT_SEQ_NO
SHA-256: $sha256Hash
>>FileTypeNum: $(Get-FileTypeNum $fileNameGUID)
>>FullPath: $fullPath

"@
                    $DIPFileStream.Write($dipEntry)

                    # Build index entry
                    $indexEntry = "$OBJ_ATT_UNID|$(Format-PadDate $OBJ_ATT_DT)|$fileNameGUID|$fullPath|$OBJ_ATT_USER_ID|$($OBJ_ATT_DSCR -replace "`r`n", " ")|$DOC_ID|$DOC_DEPT_CD|$DOC_ID|JV|$DOC_CD|$DOC_VERS_NO|$sha256Hash"
                    $IndexFileStream.WriteLine($indexEntry)
                }
            } elseif ($existsInOnBase -and $null -ne $blobData) {
                WriteLog "JVA $DOC_DEPT_CD $DOC_ID [v$DOC_VERS_NO] Attachment already in OnBase: [$OBJ_ATT_UNID] $fileName"
            } elseif ($null -eq $blobData) {
                WriteLog "DELETED ATTACHMENT [$OBJ_ATT_UNID] $fileName"
            }
        }

        $reader.Close()

        if ($fileCount -eq 0) {
            WriteLog "* No attachments found for JVA $DOC_DEPT_CD $DOC_ID - $DOC_VERS_NO"
        }

    } catch {
        WriteLog "Error retrieving attachments: $_"
    }
}

#endregion


#region Main Processing

$env:ENVIRONMENT = "DEV"  

$tmp = $ErrorActionPreference
$ErrorActionPreference = "Stop"

if ($env:ENVIRONMENT -eq "PROD") {

}else{
	function  WriteLog {
		param (
			[String]$message
		)
		Write-Host "Log: $message"
	}
}


try {
	CreateGlobalVariables $PSCommandPath $PSScriptRoot
    
    $SID = "ERP19PRO"
    $OutputPath = "N:\Projects\35442 - JVA DIP\ProcessAttachments\Attachments"
    $CheckDuplicates = $true

    WriteLog("Log Files Named")

    WriteLog "Starting JVA Attachment Processing"
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

    # Open DIP and Index files
    $dipFilePath = Join-Path $OutputPath "!JVA_attachment_indexes_DIP.txt"
    $indexFilePath = Join-Path $OutputPath "!JVA_indexes.txt"

    $dipFileStream = New-Object System.IO.StreamWriter($dipFilePath, $false, [System.Text.Encoding]::UTF8)
    $indexFileStream = New-Object System.IO.StreamWriter($indexFilePath, $false, [System.Text.Encoding]::UTF8)

    # Write index header
    $indexFileStream.WriteLine("OBJ_ATT_UNID|ATT_DATE|GUID_FILENAME|FULL_PATH|USER_ID|DESCRIPTION|DOC_ID|DEPT_CD|DOC_ID_OUT|DOC_TYP|DOC_CD|VERS_NO|SHA256")

    WriteLog "DIP file: $dipFilePath"
    WriteLog "Index file: $indexFilePath"

    # Query for JVA documents with attachments
    $jvaQuery = @"
SELECT DOC_CD, DOC_DEPT_CD, DOC_ID, DOC_VERS_NO, OBJ_ATT_PG_UNID
FROM O_FINPROD.JV_DOC_HDR
WHERE OBJ_ATT_PG_UNID IS NOT NULL AND OBJ_ATT_PG_TOT > 0
ORDER BY DOC_ID, DOC_VERS_NO
"@

    WriteLog "Querying for JVA documents..."

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

        WriteLog ""
        WriteLog "Processing JVA #$docCount : $DOC_CD $DOC_DEPT_CD $DOC_ID v$DOC_VERS_NO"

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
            -DIPFileStream $dipFileStream `
            -IndexFileStream $indexFileStream
    }

    $reader.Close()

    WriteLog ""
    WriteLog "========================================="
    WriteLog "Processing Complete"
    WriteLog "Total JVA documents processed: $docCount"
    WriteLog "Total attachments found: $script:JVAattachCount"
    WriteLog "Total attachments saved: $script:JVAattachTotal"
    WriteLog "========================================="

    # Write DIP footer
    $dipFileStream.Write("END:")

    # Clean up
    $dipFileStream.Close()
    $indexFileStream.Close()
    $erpConnection.Close()
    $onbaseConnection.Close()

    WriteLog "Script completed successfully"

} catch {
	WriteLog ($_ | Out-String).Trim()
	exit 1
} finally {
	$ErrorActionPreference = $tmp
}
#endregion