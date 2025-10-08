#Requires -Version 5.1
<#
.SYNOPSIS
    Create DIP files for JVA attachments from downloaded files
    
.DESCRIPTION
    Reads the attachment index file created by Download_JVA_Attachments.ps1
    and generates DIP files for OnBase import
    
.EXAMPLE
    .\Create_JVA_DIP_Files.ps1
#>

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
$script:ProcessDate = Get-Date -Format "MM/dd/yyyy"

#region Helper Functions

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

#endregion

#region Main Processing

try {
    $AttachmentsPath = "N:\Projects\35442 - JVA DIP\ProcessAttachments\Attachments"
    $IndexFilePath = Join-Path $AttachmentsPath "!JVA_attachment_index.txt"
    $DIPFilePath = Join-Path $AttachmentsPath "!JVA_attachment_indexes_DIP.txt"

    WriteLog "Starting DIP File Creation"
    WriteLog "Attachments Path: $AttachmentsPath"
    WriteLog "Index File: $IndexFilePath"
    WriteLog "DIP File: $DIPFilePath"

    # Check if index file exists
    if (-not (Test-Path $IndexFilePath)) {
        WriteLog "ERROR: Index file not found. Please run Download_JVA_Attachments.ps1 first."
        exit 1
    }

    # Open DIP file for writing
    $dipFileStream = New-Object System.IO.StreamWriter($DIPFilePath, $false, [System.Text.Encoding]::UTF8)

    # Read index file
    $indexLines = Get-Content $IndexFilePath
    $headerLine = $indexLines[0]
    $dataLines = $indexLines[1..($indexLines.Length - 1)]

    WriteLog "Found $($dataLines.Length) attachments in index file"

    $dipCount = 0

    foreach ($line in $dataLines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $dipCount++
        
        # Parse the index line
        $fields = $line -split '\|'
        
        $OBJ_ATT_UNID = $fields[0]
        $ATT_DATE = $fields[1]
        $GUID_FILENAME = $fields[2]
        $FULL_PATH = $fields[3]
        $USER_ID = $fields[4]
        $DESCRIPTION = $fields[5]
        $DOC_ID = $fields[6]
        $DEPT_CD = $fields[7]
        $DOC_ID_OUT = $fields[8]
        $DOC_TYP = $fields[9]
        $DOC_CD = $fields[10]
        $VERS_NO = $fields[11]
        $SG_UNID = $fields[12]
        $SEQ_NO = $fields[13]
        $STATUS = $fields[14]
        $TYPE = $fields[15]
        $COMP_NM = $fields[16]
        $COMP_DESC = $fields[17]
        $ORIGINAL_FILENAME = $fields[18]

        # Calculate SHA256 hash
        $sha256Hash = Get-SHA256Hash -FilePath $FULL_PATH

        # Build description field
        $descriptionField = ""
        if ($DESCRIPTION.Trim().Length -gt 0) {
            $descriptionField = "Long Description: $DESCRIPTION`r`n"
        }

        # Build DIP entry
        $dipEntry = @"
BEGIN:
>>Dummy Key: Document #$dipCount
>>DocTypeName: FIN - JVA Attachments
>>DocDate: $script:ProcessDate
Journal Voucher #: $DOC_ID
Advantage Attachment ID: $OBJ_ATT_UNID
Attachment Date: $ATT_DATE
${descriptionField}Filename: $ORIGINAL_FILENAME
>>Dummy Key: hidden keywords begin here
Doc ID: $DOC_ID
Version #: $VERS_NO
Department #: $DEPT_CD
Advantage Doc Type: $DOC_TYP
Advantage Doc Code: $DOC_CD
GUID File Name: $GUID_FILENAME
Advantage Attachment Primary Group ID: $OBJ_ATT_UNID
Advantage Attachment Secondary Group ID: $SG_UNID
Advantage Attachment User: $USER_ID
Advantage Attachment Status: $STATUS
Advantage Attachment Type: $TYPE
Advantage Attachment Component Name: $COMP_NM
Advantage Attachment Component Context: $COMP_DESC
Advantage Attachment Sequence #: $SEQ_NO
SHA-256: $sha256Hash
>>FileTypeNum: $(Get-FileTypeNum $GUID_FILENAME)
>>FullPath: $FULL_PATH

"@

        # Write to DIP file
        $dipFileStream.Write($dipEntry)

        if ($dipCount % 10 -eq 0) {
            WriteLog "Processed $dipCount attachments..."
        }
    }

    # Write DIP footer
    $dipFileStream.Write("END:")

    # Clean up
    $dipFileStream.Close()

    WriteLog ""
    WriteLog "========================================="
    WriteLog "DIP File Creation Complete"
    WriteLog "Total DIP entries created: $dipCount"
    WriteLog "DIP file: $DIPFilePath"
    WriteLog "========================================="

    WriteLog "DIP file creation completed successfully"

} catch {
    WriteLog "ERROR: $_"
    WriteLog $_.ScriptStackTrace
    exit 1
}

#endregion

